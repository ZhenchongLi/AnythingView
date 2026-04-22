import Cocoa

/// Renders image files using macOS native NSImageView.
class ImageRenderer: ViewerRenderer {
    static let supportedExtensions: Set<String> = [
        "png", "jpg", "jpeg", "gif", "webp", "tiff", "tif", "bmp", "ico", "heic", "heif", "svg",
    ]

    private let scrollView: NSScrollView
    private let imageView: NSImageView

    var view: NSView { scrollView }

    init() {
        scrollView = NSScrollView(frame: .zero)
        scrollView.autoresizingMask = [.width, .height]
        scrollView.hasVerticalScroller = true
        scrollView.hasHorizontalScroller = true
        scrollView.borderType = .noBorder
        scrollView.backgroundColor = .windowBackgroundColor
        scrollView.drawsBackground = true

        imageView = NSImageView(frame: .zero)
        imageView.imageScaling = .scaleProportionallyUpOrDown
        imageView.imageAlignment = .alignCenter
        imageView.autoresizingMask = [.width, .height]

        scrollView.documentView = imageView
    }

    func load(filePath: String) {
        guard let image = NSImage(contentsOfFile: filePath) else {
            DispatchQueue.main.async {
                let alert = NSAlert()
                alert.messageText = "Error"
                alert.informativeText = "Failed to load image: \(filePath)"
                alert.alertStyle = .critical
                alert.runModal()
            }
            return
        }
        DispatchQueue.main.async { [weak self] in
            guard let self else { return }
            self.imageView.image = image
            // Size the image view to the image's natural size so scroll view works
            let size = image.size
            self.imageView.frame = NSRect(origin: .zero, size: size)
        }
    }

    func setZoom(_ level: CGFloat) {
        guard let image = imageView.image else { return }
        let size = image.size
        let scaled = NSSize(width: size.width * level, height: size.height * level)
        imageView.frame = NSRect(origin: .zero, size: scaled)
    }
}
