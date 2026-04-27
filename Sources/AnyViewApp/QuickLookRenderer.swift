import Cocoa
import PDFKit
import Quartz

/// Renders files using macOS native Quick Look preview.
/// Handles pptx, xlsx, keynote, numbers, and other Quick Look-supported formats.
/// For pptx/ppt, also supports a fidelity (LibreOffice → PDF) toggle.
class QuickLookRenderer: ViewerRenderer, SupportsFidelity {
    static let supportedExtensions: Set<String> = [
        // Office / iWork (xlsx/xls now in WebRenderer via SheetJS)
        "pptx", "ppt",
        "key", "numbers", "pages",
        // Audio
        "mp3", "m4a", "wav", "flac", "aac", "aiff",
        // Video
        "mp4", "mov", "m4v", "avi",
        // 3D models
        "stl", "obj", "usdz", "usd", "dae",
        // Fonts
        "ttf", "otf", "ttc",
        // Communication
        "vcf", "ics",
    ]

    private let containerView: NSView
    private let previewView: QLPreviewView
    private let pdfView: PDFView
    private let progressIndicator: NSProgressIndicator
    private let statusLabel: NSTextField

    private var currentFilePath: String?
    private var fidelityModePreferred = false
    private var fidelityShowingPdf = false
    private var fidelityConversionID: UUID?
    private var zoomLevel: CGFloat = 1.0

    var view: NSView { containerView }

    private var fileExtension: String {
        guard let fp = currentFilePath else { return "" }
        return URL(fileURLWithPath: fp).pathExtension.lowercased()
    }

    init() {
        containerView = NSView(frame: .zero)

        previewView = QLPreviewView(frame: .zero, style: .normal)!
        previewView.autoresizingMask = [.width, .height]

        pdfView = PDFView(frame: .zero)
        pdfView.autoresizingMask = [.width, .height]
        pdfView.autoScales = true
        pdfView.displayMode = .singlePageContinuous
        pdfView.isHidden = true

        progressIndicator = NSProgressIndicator(frame: NSRect(x: 0, y: 0, width: 24, height: 24))
        progressIndicator.style = .spinning
        progressIndicator.controlSize = .small
        progressIndicator.isIndeterminate = true
        progressIndicator.isDisplayedWhenStopped = false
        progressIndicator.autoresizingMask = [.minXMargin, .maxXMargin, .minYMargin, .maxYMargin]

        statusLabel = NSTextField(frame: NSRect(x: 0, y: 0, width: 200, height: 18))
        statusLabel.isEditable = false
        statusLabel.isSelectable = false
        statusLabel.isBezeled = false
        statusLabel.drawsBackground = false
        statusLabel.alignment = .center
        statusLabel.font = NSFont.systemFont(ofSize: 12)
        statusLabel.textColor = .secondaryLabelColor
        statusLabel.stringValue = ""
        statusLabel.isHidden = true
        statusLabel.autoresizingMask = [.minXMargin, .maxXMargin, .minYMargin, .maxYMargin]

        containerView.addSubview(previewView)
        containerView.addSubview(pdfView)
        containerView.addSubview(progressIndicator)
        containerView.addSubview(statusLabel)
    }

    func load(filePath: String) {
        currentFilePath = filePath
        // Reset display state, but keep the user's fidelity preference sticky.
        fidelityShowingPdf = false
        fidelityConversionID = nil
        pdfView.document = nil
        pdfView.isHidden = true
        previewView.isHidden = false
        hideOverlay()

        let url = URL(fileURLWithPath: filePath)
        DispatchQueue.main.async { [weak self] in
            self?.previewView.previewItem = url as QLPreviewItem
        }

        if fidelityModePreferred && canEnterFidelityMode {
            DispatchQueue.main.async { [weak self] in self?.beginConversion() }
        }
    }

    func setZoom(_ level: CGFloat) {
        zoomLevel = level
        if fidelityShowingPdf {
            pdfView.scaleFactor = level
        }
        // QLPreviewView manages its own zoom internally
    }

    // MARK: - SupportsFidelity

    var canEnterFidelityMode: Bool {
        LibreOfficeFidelity.supportedExtensions.contains(fileExtension)
    }

    func setFidelityMode(_ on: Bool, completion: @escaping (Result<Void, FidelityError>) -> Void) {
        guard canEnterFidelityMode else {
            completion(.failure(.unsupportedExtension))
            return
        }
        fidelityModePreferred = on
        if on {
            beginConversion(completion: completion)
        } else {
            fidelityShowingPdf = false
            fidelityConversionID = nil
            pdfView.document = nil
            pdfView.isHidden = true
            previewView.isHidden = false
            hideOverlay()
            // Re-prime QL with the file again so it picks up cleanly.
            if let fp = currentFilePath {
                let url = URL(fileURLWithPath: fp)
                previewView.previewItem = url as QLPreviewItem
            }
            completion(.success(()))
        }
    }

    private func beginConversion(completion: ((Result<Void, FidelityError>) -> Void)? = nil) {
        guard let filePath = currentFilePath else {
            completion?(.failure(.unsupportedExtension))
            return
        }
        guard LibreOfficeCLI.findSoffice() != nil else {
            fidelityModePreferred = false
            completion?(.failure(.sofficeNotFound))
            return
        }

        let id = UUID()
        fidelityConversionID = id
        showOverlay("生成保真预览…")

        DispatchQueue.global(qos: .userInitiated).async { [weak self] in
            let result: Result<String, FidelityError>
            do {
                let pdfPath = try LibreOfficeFidelity.preparePDF(for: filePath)
                result = .success(pdfPath)
            } catch let err as FidelityError {
                result = .failure(err)
            } catch {
                result = .failure(.conversionFailed(error.localizedDescription))
            }

            DispatchQueue.main.async {
                guard let self, self.fidelityConversionID == id else { return }
                self.fidelityConversionID = nil
                self.hideOverlay()
                switch result {
                case .success(let pdf):
                    self.showPdf(at: pdf)
                    completion?(.success(()))
                case .failure(let err):
                    self.fidelityModePreferred = false
                    completion?(.failure(err))
                }
            }
        }
    }

    private func showPdf(at path: String) {
        guard let doc = PDFDocument(url: URL(fileURLWithPath: path)) else {
            fidelityModePreferred = false
            return
        }
        fidelityShowingPdf = true
        pdfView.document = doc
        pdfView.scaleFactor = zoomLevel
        previewView.isHidden = true
        pdfView.isHidden = false
    }

    private func showOverlay(_ message: String) {
        statusLabel.stringValue = message
        statusLabel.isHidden = false
        progressIndicator.startAnimation(nil)
        positionOverlay()
    }

    private func hideOverlay() {
        progressIndicator.stopAnimation(nil)
        statusLabel.isHidden = true
    }

    private func positionOverlay() {
        let bounds = containerView.bounds
        let spinSize: CGFloat = 24
        let labelWidth: CGFloat = 200
        let labelHeight: CGFloat = 18
        let gap: CGFloat = 8
        let totalHeight = spinSize + gap + labelHeight
        let originY = (bounds.height - totalHeight) / 2
        progressIndicator.frame = NSRect(
            x: (bounds.width - spinSize) / 2,
            y: originY + labelHeight + gap,
            width: spinSize, height: spinSize
        )
        statusLabel.frame = NSRect(
            x: (bounds.width - labelWidth) / 2,
            y: originY,
            width: labelWidth, height: labelHeight
        )
    }
}
