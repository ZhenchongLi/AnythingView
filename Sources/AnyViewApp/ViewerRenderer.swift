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

/// Errors surfaced by fidelity (LibreOffice → PDF) conversion.
enum FidelityError: LocalizedError {
    case sofficeNotFound
    case noSourceDocx
    case conversionFailed(String)
    case unsupportedExtension

    var errorDescription: String? {
        switch self {
        case .sofficeNotFound: return "LibreOffice (soffice) not found."
        case .noSourceDocx: return "No source.docx found inside the package."
        case .conversionFailed(let msg): return "LibreOffice conversion failed: \(msg)"
        case .unsupportedExtension: return "Fidelity mode does not support this file type."
        }
    }
}

/// Renderers that can swap their normal display for a high-fidelity
/// LibreOffice → PDF render of the same file.
protocol SupportsFidelity: AnyObject {
    var canEnterFidelityMode: Bool { get }
    func setFidelityMode(_ on: Bool, completion: @escaping (Result<Void, FidelityError>) -> Void)
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
