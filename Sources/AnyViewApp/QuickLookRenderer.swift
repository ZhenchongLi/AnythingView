import Cocoa
import Quartz

/// Renders files using macOS native Quick Look preview.
/// Handles pptx, xlsx, keynote, numbers, and other Quick Look-supported formats.
class QuickLookRenderer: ViewerRenderer {
    static let supportedExtensions: Set<String> = [
        // Office / iWork
        "pptx", "ppt", "xlsx", "xls",
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
