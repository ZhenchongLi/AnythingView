import Cocoa
import UniformTypeIdentifiers

class AppDelegate: NSObject, NSApplicationDelegate, NSMenuItemValidation {
    private var windows: [ViewerWindowController] = []
    private var hasOpenedFile = false

    func applicationDidFinishLaunching(_ notification: Notification) {
        // If no file was opened via double-click, show a file picker
        DispatchQueue.main.async { [weak self] in
            guard let self, !self.hasOpenedFile else { return }
            self.showOpenPanel()
        }
    }

    func application(_ sender: NSApplication, openFile filename: String) -> Bool {
        hasOpenedFile = true
        openDocument(at: filename)
        return true
    }

    func application(_ sender: NSApplication, openFiles filenames: [String]) {
        hasOpenedFile = true
        for filename in filenames {
            openDocument(at: filename)
        }
    }

    func applicationShouldTerminateAfterLastWindowClosed(_ sender: NSApplication) -> Bool {
        return true
    }

    func applicationShouldHandleReopen(_ sender: NSApplication, hasVisibleWindows flag: Bool) -> Bool {
        if !flag {
            showOpenPanel()
        }
        return true
    }

    // MARK: - Appearance

    private var currentAppearance: Int = 0  // 0=system, 1=light, 2=dark

    @objc func setAppearanceSystem(_ sender: Any?) { setAppearance(0) }
    @objc func setAppearanceLight(_ sender: Any?) { setAppearance(1) }
    @objc func setAppearanceDark(_ sender: Any?) { setAppearance(2) }

    private func setAppearance(_ mode: Int) {
        currentAppearance = mode
        switch mode {
        case 1: NSApp.appearance = NSAppearance(named: .aqua)
        case 2: NSApp.appearance = NSAppearance(named: .darkAqua)
        default: NSApp.appearance = nil  // follow system
        }
    }

    // MARK: - Zoom

    @objc func zoomIn(_ sender: Any?) { activeController()?.zoomIn(sender) }
    @objc func zoomOut(_ sender: Any?) { activeController()?.zoomOut(sender) }
    @objc func actualSize(_ sender: Any?) { activeController()?.actualSize(sender) }

    // MARK: - Reload

    @objc func reload(_ sender: Any?) { activeController()?.reload(sender) }

    private func activeController() -> ViewerWindowController? {
        guard let key = NSApp.keyWindow else { return nil }
        return windows.first { $0.window === key }
    }

    func validateMenuItem(_ menuItem: NSMenuItem) -> Bool {
        if menuItem.action == #selector(setAppearanceSystem(_:)) ||
           menuItem.action == #selector(setAppearanceLight(_:)) ||
           menuItem.action == #selector(setAppearanceDark(_:)) {
            menuItem.state = menuItem.tag == currentAppearance ? .on : .off
            return true
        }
        if menuItem.action == #selector(zoomIn(_:)) ||
           menuItem.action == #selector(zoomOut(_:)) ||
           menuItem.action == #selector(actualSize(_:)) ||
           menuItem.action == #selector(reload(_:)) {
            return activeController() != nil
        }
        return true
    }

    // MARK: - Open

    @objc func openDocument(_ sender: Any?) {
        showOpenPanel()
    }

    private func openDocument(at path: String) {
        let resolved = (path as NSString).standardizingPath
        let url = URL(fileURLWithPath: resolved)
        let ext = url.pathExtension.lowercased()
        guard RendererFactory.allSupportedExtensions.contains(ext) else {
            showError("Unsupported file type: .\(ext)")
            return
        }

        // Activate existing tab if already open
        if let existing = windows.first(where: { $0.filePath == resolved }) {
            existing.activate()
            return
        }

        let controller = ViewerWindowController(filePath: resolved)
        controller.onClose = { [weak self] ctrl in
            self?.windows.removeAll { $0 === ctrl }
        }
        controller.onOpenFiles = { [weak self] paths in
            for path in paths {
                self?.openDocument(at: path)
            }
        }
        windows.append(controller)
        controller.showWindow(nil)
    }

    private func showOpenPanel() {
        let panel = NSOpenPanel()
        panel.title = "Open Document"
        panel.allowsMultipleSelection = true
        panel.canChooseDirectories = false
        panel.canChooseFiles = true
        var types: [UTType] = []
        for ext in RendererFactory.allSupportedExtensions {
            if let t = UTType(filenameExtension: ext) { types.append(t) }
        }
        panel.allowedContentTypes = types.isEmpty ? [.data] : types

        if panel.runModal() == .OK {
            for url in panel.urls {
                openDocument(at: url.path)
            }
        }
    }

    private func showError(_ message: String) {
        let alert = NSAlert()
        alert.messageText = "Error"
        alert.informativeText = message
        alert.alertStyle = .critical
        alert.runModal()
    }
}
