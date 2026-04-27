import Cocoa

private extension NSToolbarItem.Identifier {
    static let appearanceToggle = NSToolbarItem.Identifier("appearanceToggle")
    static let zoomOut = NSToolbarItem.Identifier("zoomOut")
    static let zoomReset = NSToolbarItem.Identifier("zoomReset")
    static let zoomIn = NSToolbarItem.Identifier("zoomIn")
    static let fidelityToggle = NSToolbarItem.Identifier("fidelityToggle")
}

class ViewerWindowController: NSObject, NSWindowDelegate, NSToolbarDelegate {

    static let minZoom: CGFloat = 0.5
    static let maxZoom: CGFloat = 3.0
    static let zoomStep: CGFloat = 0.1

    static let reloadDebounceInterval: DispatchTimeInterval = .milliseconds(250)

    let filePath: String
    var onClose: ((ViewerWindowController) -> Void)?
    var onOpenFiles: (([String]) -> Void)?

    private(set) var window: NSWindow?
    private var renderer: ViewerRenderer?
    private var zoomLevel: CGFloat = 1.0
    private weak var zoomLabelButton: NSButton?

    private var findBar: FindBarView?
    private var rendererContainer: NSView?

    private var watcherSource: DispatchSourceFileSystemObject?
    private var reloadDebounceItem: DispatchWorkItem?
    private let reloadQueue = DispatchQueue(label: "com.anyview.reload", qos: .userInitiated)

    private weak var fidelityToolbarItem: NSToolbarItem?
    private var fidelityOn: Bool = false

    private var fileExtension: String {
        URL(fileURLWithPath: filePath).pathExtension.lowercased()
    }

    init(filePath: String) {
        self.filePath = filePath
        super.init()
    }

    deinit {
        reloadDebounceItem?.cancel()
        stopWatching()
    }

    func showWindow(_ sender: Any?) {
        let filename = URL(fileURLWithPath: filePath).lastPathComponent

        let screen = NSScreen.main?.visibleFrame ?? NSRect(x: 0, y: 0, width: 900, height: 900)
        let width = min(900.0, screen.width * 0.8)
        let height = min(1100.0, screen.height * 0.9)
        let contentRect = NSRect(x: 0, y: 0, width: width, height: height)
        let styleMask: NSWindow.StyleMask = [.titled, .closable, .miniaturizable, .resizable]
        let win = NSWindow(contentRect: contentRect, styleMask: styleMask,
                           backing: .buffered, defer: false)
        win.isReleasedWhenClosed = false
        win.title = filename
        win.delegate = self
        win.tabbingMode = .preferred
        win.tabbingIdentifier = "AnyView"

        let r = RendererFactory.renderer(for: fileExtension)
        self.renderer = r

        let dropTarget = DropTargetView(frame: win.contentView?.bounds ?? .zero)
        dropTarget.autoresizingMask = [.width, .height]
        dropTarget.onDrop = { [weak self] paths in
            self?.onOpenFiles?(paths)
        }

        // Renderer lives inside a container so the find bar can be inserted above it.
        let container = NSView(frame: dropTarget.bounds)
        container.autoresizingMask = [.width, .height]
        r.view.frame = container.bounds
        r.view.autoresizingMask = [.width, .height]
        container.addSubview(r.view)
        dropTarget.addSubview(container)
        self.rendererContainer = container
        win.contentView = dropTarget

        let toolbar = NSToolbar(identifier: "AnyViewToolbar")
        toolbar.delegate = self
        toolbar.displayMode = .iconOnly
        win.toolbar = toolbar
        win.titleVisibility = .visible

        win.center()
        win.makeKeyAndOrderFront(nil)
        self.window = win

        startWatching()
        reloadQueue.async { [weak self] in
            guard let self else { return }
            self.renderer?.load(filePath: self.filePath)
        }
    }

    func activate() {
        window?.makeKeyAndOrderFront(nil)
    }

    // MARK: - Reload

    @objc func reload(_ sender: Any?) {
        performReload()
    }

    private func performReload() {
        reloadDebounceItem?.cancel()
        stopWatching()
        startWatching()
        reloadQueue.async { [weak self] in
            guard let self else { return }
            self.renderer?.load(filePath: self.filePath)
        }
    }

    // MARK: - File Watching

    private func startWatching() {
        let fd = open(filePath, O_EVTONLY)
        guard fd >= 0 else { return }
        let source = DispatchSource.makeFileSystemObjectSource(
            fileDescriptor: fd,
            eventMask: [.write, .delete, .rename, .extend],
            queue: DispatchQueue.main
        )
        source.setEventHandler { [weak self] in
            self?.scheduleReload()
        }
        source.setCancelHandler { close(fd) }
        watcherSource = source
        source.resume()
    }

    private func stopWatching() {
        watcherSource?.cancel()
        watcherSource = nil
    }

    private func scheduleReload() {
        reloadDebounceItem?.cancel()
        let item = DispatchWorkItem { [weak self] in
            self?.performReload()
        }
        reloadDebounceItem = item
        DispatchQueue.main.asyncAfter(deadline: .now() + Self.reloadDebounceInterval, execute: item)
    }

    // MARK: - Fidelity Mode

    @objc func toggleFidelity(_ sender: Any?) {
        guard let renderer = renderer as? SupportsFidelity else { return }
        let newState = !fidelityOn
        if newState && LibreOfficeCLI.findSoffice() == nil {
            promptInstallLibreOffice { [weak self] proceed in
                guard let self else { return }
                if proceed {
                    self.beginInstallLibreOffice()
                }
            }
            return
        }
        renderer.setFidelityMode(newState) { [weak self] result in
            DispatchQueue.main.async {
                guard let self else { return }
                switch result {
                case .success:
                    self.fidelityOn = newState
                    self.updateFidelityToolbarIcon()
                case .failure(let err):
                    self.fidelityOn = false
                    self.updateFidelityToolbarIcon()
                    self.showFidelityError(err)
                }
            }
        }
    }

    private func updateFidelityToolbarIcon() {
        let symbol = fidelityOn ? "doc.richtext.fill" : "doc.richtext"
        fidelityToolbarItem?.image = NSImage(
            systemSymbolName: symbol,
            accessibilityDescription: "Fidelity Mode"
        )
    }

    private func promptInstallLibreOffice(completion: @escaping (Bool) -> Void) {
        let alert = NSAlert()
        alert.messageText = "需要 LibreOffice 才能开启保真模式"
        alert.informativeText = "保真模式用 LibreOffice 把文档渲染成 PDF（公文红头、复杂表格、字段都对得上）。\n首次使用需要下载并安装到 AnyView 应用数据目录，约 300 MB。"
        alert.alertStyle = .informational
        alert.addButton(withTitle: "下载并安装")
        alert.addButton(withTitle: "取消")
        if let win = window {
            alert.beginSheetModal(for: win) { resp in
                completion(resp == .alertFirstButtonReturn)
            }
        } else {
            completion(alert.runModal() == .alertFirstButtonReturn)
        }
    }

    private func beginInstallLibreOffice() {
        let installer = LibreOfficeInstaller(attachWindow: window)
        installer.runWithSheet { [weak self] result in
            guard let self else { return }
            switch result {
            case .success:
                // Re-trigger the toggle now that soffice exists.
                self.toggleFidelity(nil)
            case .failure(let err):
                if case .canceled = err { return }
                let alert = NSAlert()
                alert.messageText = "LibreOffice 安装失败"
                alert.informativeText = err.errorDescription ?? "未知错误"
                alert.alertStyle = .warning
                alert.addButton(withTitle: "好")
                if let win = self.window {
                    alert.beginSheetModal(for: win, completionHandler: nil)
                } else {
                    alert.runModal()
                }
            }
        }
    }

    private func showFidelityError(_ err: FidelityError) {
        let alert = NSAlert()
        alert.messageText = "保真模式出错"
        alert.informativeText = err.errorDescription ?? "未知错误"
        alert.alertStyle = .warning
        alert.addButton(withTitle: "好")
        if let win = window {
            alert.beginSheetModal(for: win, completionHandler: nil)
        } else {
            alert.runModal()
        }
    }

    // MARK: - Appearance Toggle

    @objc private func toggleAppearance(_ sender: Any?) {
        let isDark = window?.effectiveAppearance.bestMatch(from: [.darkAqua, .aqua]) == .darkAqua
        NSApp.appearance = isDark ? NSAppearance(named: .aqua) : NSAppearance(named: .darkAqua)
        if let item = window?.toolbar?.items.first(where: { $0.itemIdentifier == .appearanceToggle }) {
            item.image = NSImage(
                systemSymbolName: isDark ? "moon.circle" : "sun.max.circle",
                accessibilityDescription: "Toggle appearance"
            )
        }
    }

    // MARK: - Zoom

    @objc func zoomIn(_ sender: Any?) { setZoom(zoomLevel + Self.zoomStep) }
    @objc func zoomOut(_ sender: Any?) { setZoom(zoomLevel - Self.zoomStep) }
    @objc func actualSize(_ sender: Any?) { setZoom(1.0) }

    private var zoomLabelText: String { "\(Int((zoomLevel * 100).rounded()))%" }

    private func setZoom(_ value: CGFloat) {
        let snapped = (value * 10).rounded() / 10
        zoomLevel = min(max(snapped, Self.minZoom), Self.maxZoom)
        renderer?.setZoom(zoomLevel)
        zoomLabelButton?.title = zoomLabelText
    }

    // MARK: - NSToolbarDelegate

    private var fidelityCapable: Bool {
        LibreOfficeFidelity.supportedExtensions.contains(fileExtension)
    }

    func toolbarDefaultItemIdentifiers(_ toolbar: NSToolbar) -> [NSToolbarItem.Identifier] {
        var ids: [NSToolbarItem.Identifier] = [.zoomOut, .zoomReset, .zoomIn, .flexibleSpace]
        if fidelityCapable { ids.append(.fidelityToggle) }
        ids.append(.appearanceToggle)
        return ids
    }

    func toolbarAllowedItemIdentifiers(_ toolbar: NSToolbar) -> [NSToolbarItem.Identifier] {
        [.zoomOut, .zoomReset, .zoomIn, .flexibleSpace, .fidelityToggle, .appearanceToggle]
    }

    func toolbar(_ toolbar: NSToolbar, itemForItemIdentifier itemIdentifier: NSToolbarItem.Identifier, willBeInsertedIntoToolbar flag: Bool) -> NSToolbarItem? {
        if itemIdentifier == .appearanceToggle {
            let item = NSToolbarItem(itemIdentifier: .appearanceToggle)
            let isDark = NSApp.effectiveAppearance.bestMatch(from: [.darkAqua, .aqua]) == .darkAqua
            item.image = NSImage(
                systemSymbolName: isDark ? "sun.max.circle" : "moon.circle",
                accessibilityDescription: "Toggle appearance"
            )
            item.label = "Appearance"
            item.toolTip = "Toggle Dark / Light"
            item.target = self
            item.action = #selector(toggleAppearance(_:))
            return item
        }
        if itemIdentifier == .zoomOut {
            let item = NSToolbarItem(itemIdentifier: .zoomOut)
            item.image = NSImage(systemSymbolName: "minus.magnifyingglass", accessibilityDescription: "Zoom Out")
            item.label = "Zoom Out"
            item.toolTip = "Zoom Out (⌘−)"
            item.target = self
            item.action = #selector(zoomOut(_:))
            return item
        }
        if itemIdentifier == .zoomIn {
            let item = NSToolbarItem(itemIdentifier: .zoomIn)
            item.image = NSImage(systemSymbolName: "plus.magnifyingglass", accessibilityDescription: "Zoom In")
            item.label = "Zoom In"
            item.toolTip = "Zoom In (⌘+)"
            item.target = self
            item.action = #selector(zoomIn(_:))
            return item
        }
        if itemIdentifier == .zoomReset {
            let button = NSButton(title: zoomLabelText, target: self, action: #selector(actualSize(_:)))
            button.bezelStyle = .texturedRounded
            button.font = NSFont.systemFont(ofSize: 12, weight: .medium)
            button.translatesAutoresizingMaskIntoConstraints = false
            button.widthAnchor.constraint(equalToConstant: 56).isActive = true
            self.zoomLabelButton = button
            let item = NSToolbarItem(itemIdentifier: .zoomReset)
            item.view = button
            item.label = "Zoom"
            item.toolTip = "Reset Zoom (⌘0)"
            return item
        }
        if itemIdentifier == .fidelityToggle {
            let item = NSToolbarItem(itemIdentifier: .fidelityToggle)
            item.image = NSImage(
                systemSymbolName: "doc.richtext",
                accessibilityDescription: "Fidelity Mode"
            )
            item.label = "保真"
            item.toolTip = "切换到 LibreOffice 高保真渲染"
            item.target = self
            item.action = #selector(toggleFidelity(_:))
            self.fidelityToolbarItem = item
            return item
        }
        return nil
    }

    // MARK: - Find

    @objc func performFind(_ sender: Any?) {
        showFindBar()
    }

    @objc func findNext(_ sender: Any?) {
        triggerFind(forward: true)
    }

    @objc func findPrevious(_ sender: Any?) {
        triggerFind(forward: false)
    }

    var supportsFind: Bool { renderer is SupportsFind }

    private func showFindBar() {
        guard let container = rendererContainer, let parent = container.superview else { return }

        if let bar = findBar {
            bar.focusInput()
            return
        }

        let barHeight = FindBarView.preferredHeight
        let bar = FindBarView(frame: NSRect(x: 0,
                                            y: parent.bounds.height - barHeight,
                                            width: parent.bounds.width,
                                            height: barHeight))
        bar.autoresizingMask = [.width, .minYMargin]
        bar.delegate = self
        parent.addSubview(bar)

        var f = container.frame
        f.size.height = parent.bounds.height - barHeight
        container.frame = f

        findBar = bar
        bar.focusInput()
    }

    private func hideFindBar() {
        guard let bar = findBar, let container = rendererContainer, let parent = container.superview else { return }
        bar.removeFromSuperview()
        container.frame = parent.bounds
        findBar = nil
        window?.makeFirstResponder(renderer?.view)
    }

    private func triggerFind(forward: Bool) {
        guard let bar = findBar else {
            showFindBar()
            return
        }
        let q = bar.query
        guard !q.isEmpty else {
            bar.focusInput()
            return
        }
        guard let finder = renderer as? SupportsFind else {
            bar.setStatus("Not supported", isError: true)
            return
        }
        finder.performFind(query: q, forward: forward) { [weak bar] found in
            DispatchQueue.main.async {
                if found {
                    bar?.setStatus("")
                } else {
                    bar?.setStatus("Not found", isError: true)
                }
            }
        }
    }

    // MARK: - NSWindowDelegate

    func windowWillClose(_ notification: Notification) {
        reloadDebounceItem?.cancel()
        reloadDebounceItem = nil
        stopWatching()
        if let webRenderer = renderer as? WebRenderer {
            webRenderer.cleanup()
        }
        onClose?(self)
    }
}

extension ViewerWindowController: FindBarViewDelegate {
    func findBar(_ bar: FindBarView, didSearch query: String, forward: Bool) {
        triggerFind(forward: forward)
    }

    func findBarDidRequestClose(_ bar: FindBarView) {
        hideFindBar()
    }
}
