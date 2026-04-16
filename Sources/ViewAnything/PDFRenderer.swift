import Cocoa
import PDFKit

/// Renders PDF files using macOS native PDFView.
class PDFRenderer: ViewerRenderer {
    static let supportedExtensions: Set<String> = ["pdf"]

    private let pdfView: PDFView

    var view: NSView { pdfView }

    init() {
        pdfView = PDFView(frame: .zero)
        pdfView.autoresizingMask = [.width, .height]
        pdfView.autoScales = true
        pdfView.displayMode = .singlePageContinuous
    }

    func load(filePath: String) {
        guard let document = PDFDocument(url: URL(fileURLWithPath: filePath)) else {
            DispatchQueue.main.async {
                let alert = NSAlert()
                alert.messageText = "Error"
                alert.informativeText = "Failed to load PDF: \(filePath)"
                alert.alertStyle = .critical
                alert.runModal()
            }
            return
        }
        DispatchQueue.main.async { [weak self] in
            self?.pdfView.document = document
        }
    }

    func setZoom(_ level: CGFloat) {
        pdfView.scaleFactor = level
    }
}
