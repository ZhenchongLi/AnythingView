import Cocoa
import Quartz

/// Renders files using macOS native Quick Look preview.
/// Handles pptx, xlsx, keynote, numbers, and other Quick Look-supported formats.
class QuickLookRenderer: ViewerRenderer {
    static let supportedExtensions: Set<String> = [
        "pptx", "ppt", "xlsx", "xls",
        "key", "numbers", "pages",
    ]

    private let previewView: QLPreviewView

    var view: NSView { previewView }

    init() {
        previewView = QLPreviewView(frame: .zero, style: .normal)!
        previewView.autoresizingMask = [.width, .height]
    }

    func load(filePath: String) {
        let url = URL(fileURLWithPath: filePath)
        DispatchQueue.main.async { [weak self] in
            self?.previewView.previewItem = url as QLPreviewItem
        }
    }

    func setZoom(_ level: CGFloat) {
        // QLPreviewView manages its own zoom internally
    }
}
