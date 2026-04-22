import Cocoa

@main
struct AnyViewApp {
    // Strong reference — NSApplication.delegate is weak, so without this
    // the delegate gets released during termination and crashes.
    static let appDelegate = AppDelegate()

    static func main() {
        let app = NSApplication.shared
        app.setActivationPolicy(.regular)
        app.mainMenu = buildMainMenu()
        app.delegate = appDelegate
        app.activate(ignoringOtherApps: true)
        app.run()
    }

    static func buildMainMenu() -> NSMenu {
        let mainMenu = NSMenu()

        // App menu
        let appMenuItem = NSMenuItem()
        mainMenu.addItem(appMenuItem)
        let appMenu = NSMenu()
        appMenuItem.submenu = appMenu
        appMenu.addItem(withTitle: "About AnyView", action: #selector(NSApplication.orderFrontStandardAboutPanel(_:)), keyEquivalent: "")
        appMenu.addItem(.separator())
        appMenu.addItem(withTitle: "Hide AnyView", action: #selector(NSApplication.hide(_:)), keyEquivalent: "h")
        let hideOthers = appMenu.addItem(withTitle: "Hide Others", action: #selector(NSApplication.hideOtherApplications(_:)), keyEquivalent: "h")
        hideOthers.keyEquivalentModifierMask = [.command, .option]
        appMenu.addItem(withTitle: "Show All", action: #selector(NSApplication.unhideAllApplications(_:)), keyEquivalent: "")
        appMenu.addItem(.separator())
        appMenu.addItem(withTitle: "Quit AnyView", action: #selector(NSApplication.terminate(_:)), keyEquivalent: "q")

        // File menu
        let fileMenuItem = NSMenuItem()
        mainMenu.addItem(fileMenuItem)
        let fileMenu = NSMenu(title: "File")
        fileMenuItem.submenu = fileMenu
        fileMenu.addItem(withTitle: "Open…", action: #selector(AppDelegate.openDocument(_:)), keyEquivalent: "o")
        fileMenu.addItem(withTitle: "Close", action: #selector(NSWindow.performClose(_:)), keyEquivalent: "w")

        // Edit menu — routes ⌘C/⌘A/⌘X/⌘V to first responder (WKWebView)
        let editMenuItem = NSMenuItem()
        mainMenu.addItem(editMenuItem)
        let editMenu = NSMenu(title: "Edit")
        editMenuItem.submenu = editMenu
        editMenu.addItem(withTitle: "Cut", action: #selector(NSText.cut(_:)), keyEquivalent: "x")
        editMenu.addItem(withTitle: "Copy", action: #selector(NSText.copy(_:)), keyEquivalent: "c")
        editMenu.addItem(withTitle: "Paste", action: #selector(NSText.paste(_:)), keyEquivalent: "v")
        editMenu.addItem(.separator())
        editMenu.addItem(withTitle: "Select All", action: #selector(NSText.selectAll(_:)), keyEquivalent: "a")
        editMenu.addItem(.separator())

        let findMenu = NSMenu(title: "Find")
        let findItem = editMenu.addItem(withTitle: "Find", action: nil, keyEquivalent: "")
        editMenu.setSubmenu(findMenu, for: findItem)
        findMenu.addItem(withTitle: "Find…", action: #selector(AppDelegate.performFind(_:)), keyEquivalent: "f")
        let findNextItem = findMenu.addItem(withTitle: "Find Next", action: #selector(AppDelegate.findNext(_:)), keyEquivalent: "g")
        findNextItem.keyEquivalentModifierMask = [.command]
        let findPrevItem = findMenu.addItem(withTitle: "Find Previous", action: #selector(AppDelegate.findPrevious(_:)), keyEquivalent: "g")
        findPrevItem.keyEquivalentModifierMask = [.command, .shift]

        // View menu
        let viewMenuItem = NSMenuItem()
        mainMenu.addItem(viewMenuItem)
        let viewMenu = NSMenu(title: "View")
        viewMenuItem.submenu = viewMenu

        let reloadItem = viewMenu.addItem(withTitle: "Reload", action: #selector(AppDelegate.reload(_:)), keyEquivalent: "r")
        reloadItem.keyEquivalentModifierMask = [.command]
        viewMenu.addItem(.separator())

        let zoomInItem = viewMenu.addItem(withTitle: "Zoom In", action: #selector(AppDelegate.zoomIn(_:)), keyEquivalent: "+")
        zoomInItem.keyEquivalentModifierMask = [.command]
        // Hidden duplicate so Cmd+= (without shift) also triggers Zoom In
        let zoomInAlt = NSMenuItem(title: "Zoom In", action: #selector(AppDelegate.zoomIn(_:)), keyEquivalent: "=")
        zoomInAlt.keyEquivalentModifierMask = [.command]
        zoomInAlt.isHidden = true
        zoomInAlt.allowsKeyEquivalentWhenHidden = true
        viewMenu.addItem(zoomInAlt)
        let zoomOutItem = viewMenu.addItem(withTitle: "Zoom Out", action: #selector(AppDelegate.zoomOut(_:)), keyEquivalent: "-")
        zoomOutItem.keyEquivalentModifierMask = [.command]
        let actualSizeItem = viewMenu.addItem(withTitle: "Actual Size", action: #selector(AppDelegate.actualSize(_:)), keyEquivalent: "0")
        actualSizeItem.keyEquivalentModifierMask = [.command]
        viewMenu.addItem(.separator())

        let appearanceMenu = NSMenu(title: "Appearance")
        let appearanceItem = viewMenu.addItem(withTitle: "Appearance", action: nil, keyEquivalent: "")
        viewMenu.setSubmenu(appearanceMenu, for: appearanceItem)

        let systemItem = appearanceMenu.addItem(withTitle: "System", action: #selector(AppDelegate.setAppearanceSystem(_:)), keyEquivalent: "")
        systemItem.tag = 0
        let lightItem = appearanceMenu.addItem(withTitle: "Light", action: #selector(AppDelegate.setAppearanceLight(_:)), keyEquivalent: "")
        lightItem.tag = 1
        let darkItem = appearanceMenu.addItem(withTitle: "Dark", action: #selector(AppDelegate.setAppearanceDark(_:)), keyEquivalent: "")
        darkItem.tag = 2

        // Window menu
        let windowMenuItem = NSMenuItem()
        mainMenu.addItem(windowMenuItem)
        let windowMenu = NSMenu(title: "Window")
        windowMenuItem.submenu = windowMenu
        windowMenu.addItem(withTitle: "Minimize", action: #selector(NSWindow.performMiniaturize(_:)), keyEquivalent: "m")
        windowMenu.addItem(withTitle: "Zoom", action: #selector(NSWindow.performZoom(_:)), keyEquivalent: "")
        windowMenu.addItem(.separator())
        windowMenu.addItem(withTitle: "Show All Windows", action: #selector(NSApplication.arrangeInFront(_:)), keyEquivalent: "")
        NSApp.windowsMenu = windowMenu

        return mainMenu
    }
}
