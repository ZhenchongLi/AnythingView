import Cocoa

/// Pluggable rendering backend for AnyView.
/// Each renderer owns its NSView and knows how to load/reload a specific set of file types.
protocol ViewerRenderer: AnyObject {
    /// The view to embed in the window.
    var view: NSView { get }

    /// File extensions this renderer handles.
    static var supportedExtensions: Set<String> { get }

    /// Load the file at the given path.
    func load(filePath: String)

    /// Set the zoom level (1.0 = 100%).
    func setZoom(_ level: CGFloat)
}

/// Renderers that support in-document text search.
protocol SupportsFind: AnyObject {
    func performFind(query: String, forward: Bool, completion: @escaping (Bool) -> Void)
}

/// Returns the appropriate renderer for a file extension.
enum RendererFactory {
    static func renderer(for extension: String) -> ViewerRenderer {
        let ext = `extension`.lowercased()
        if PDFRenderer.supportedExtensions.contains(ext) {
            return PDFRenderer()
        }
        if ImageRenderer.supportedExtensions.contains(ext) {
            return ImageRenderer()
        }
        if QuickLookRenderer.supportedExtensions.contains(ext) {
            return QuickLookRenderer()
        }
        // Default: web renderer handles everything else
        return WebRenderer()
    }

    static var allSupportedExtensions: Set<String> {
        PDFRenderer.supportedExtensions
            .union(ImageRenderer.supportedExtensions)
            .union(QuickLookRenderer.supportedExtensions)
            .union(WebRenderer.supportedExtensions)
    }
}
