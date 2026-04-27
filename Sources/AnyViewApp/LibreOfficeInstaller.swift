import Cocoa

/// In-app LibreOffice download + install. Pulls the official TDF DMG, mounts it,
/// copies `LibreOffice.app` into `~/Library/Application Support/AnyView/`,
/// then detaches. No external dependencies; uses only `URLSession`, `hdiutil`,
/// `xattr`, and `cp`.
final class LibreOfficeInstaller: NSObject, URLSessionDownloadDelegate {

    /// Pinned to a known-published TDF release. Bump in source when newer
    /// builds are validated. The TDF mirror has no "latest" symlink so a
    /// dynamic discovery is more brittle than just shipping a version.
    static let version = "26.2.2"

    enum InstallError: LocalizedError {
        case unsupportedArch
        case downloadFailed(String)
        case mountFailed(String)
        case copyFailed(String)
        case verificationFailed
        case canceled

        var errorDescription: String? {
            switch self {
            case .unsupportedArch: return "Unsupported CPU architecture."
            case .downloadFailed(let m): return "Download failed: \(m)"
            case .mountFailed(let m): return "DMG mount failed: \(m)"
            case .copyFailed(let m): return "Copy failed: \(m)"
            case .verificationFailed: return "Installed binary not executable."
            case .canceled: return "Canceled."
            }
        }
    }

    enum Phase {
        case downloading(received: Int64, total: Int64)
        case mounting
        case copying
        case finalizing
    }

    private let attachWindow: NSWindow?
    private var session: URLSession?
    private var task: URLSessionDownloadTask?
    private var canceled = false
    private var dmgURL: URL!
    private var sheet: InstallProgressSheet?

    private var onPhase: ((Phase) -> Void)?
    private var onFinish: ((Result<String, InstallError>) -> Void)?

    init(attachWindow: NSWindow?) {
        self.attachWindow = attachWindow
        super.init()
    }

    /// Kick off the install with a progress sheet attached to the parent window.
    /// `completion` runs on the main thread.
    func runWithSheet(completion: @escaping (Result<String, InstallError>) -> Void) {
        guard let dmgURL = Self.dmgURL() else {
            completion(.failure(.unsupportedArch))
            return
        }
        self.dmgURL = dmgURL

        let sheet = InstallProgressSheet()
        sheet.onCancel = { [weak self] in self?.cancel() }
        self.sheet = sheet

        if let win = attachWindow {
            win.beginSheet(sheet.window) { _ in }
        } else {
            sheet.window.makeKeyAndOrderFront(nil)
        }

        run(onPhase: { [weak sheet] phase in
            DispatchQueue.main.async { sheet?.update(phase) }
        }, completion: { [weak self] result in
            DispatchQueue.main.async {
                if let win = self?.attachWindow, let s = self?.sheet?.window {
                    win.endSheet(s)
                } else {
                    self?.sheet?.window.orderOut(nil)
                }
                completion(result)
            }
        })
    }

    func cancel() {
        canceled = true
        task?.cancel()
    }

    private func run(onPhase: @escaping (Phase) -> Void,
                     completion: @escaping (Result<String, InstallError>) -> Void) {
        self.onPhase = onPhase
        self.onFinish = completion

        let config = URLSessionConfiguration.default
        let session = URLSession(configuration: config, delegate: self, delegateQueue: nil)
        self.session = session
        let task = session.downloadTask(with: dmgURL)
        self.task = task
        onPhase(.downloading(received: 0, total: 0))
        task.resume()
    }

    // MARK: URLSessionDownloadDelegate

    func urlSession(_ session: URLSession, downloadTask: URLSessionDownloadTask,
                    didWriteData bytesWritten: Int64,
                    totalBytesWritten: Int64,
                    totalBytesExpectedToWrite: Int64) {
        onPhase?(.downloading(received: totalBytesWritten, total: totalBytesExpectedToWrite))
    }

    func urlSession(_ session: URLSession, downloadTask: URLSessionDownloadTask,
                    didFinishDownloadingTo location: URL) {
        // URLSession deletes `location` once this method returns; move it first.
        let staging = URL(fileURLWithPath: NSTemporaryDirectory())
            .appendingPathComponent("AnyView-LO-\(UUID().uuidString).dmg")
        do {
            try FileManager.default.moveItem(at: location, to: staging)
        } catch {
            finish(.failure(.downloadFailed(error.localizedDescription)))
            return
        }

        // Mount + copy + detach happen on a background queue so we don't block
        // delegate callbacks (and so cancel still works).
        DispatchQueue.global(qos: .userInitiated).async { [weak self] in
            guard let self else { return }
            if self.canceled { self.finish(.failure(.canceled)); return }
            self.mountCopyDetach(dmg: staging)
        }
    }

    func urlSession(_ session: URLSession, task: URLSessionTask,
                    didCompleteWithError error: Error?) {
        if let err = error as NSError?,
           err.domain == NSURLErrorDomain && err.code == NSURLErrorCancelled {
            finish(.failure(.canceled))
            return
        }
        if let err = error {
            finish(.failure(.downloadFailed(err.localizedDescription)))
        }
    }

    // MARK: Helpers

    private func mountCopyDetach(dmg: URL) {
        defer { try? FileManager.default.removeItem(at: dmg) }

        onPhase?(.mounting)
        let mountPoint: String
        do {
            mountPoint = try Self.attachDMG(at: dmg.path)
        } catch let err as InstallError {
            finish(.failure(err)); return
        } catch {
            finish(.failure(.mountFailed(error.localizedDescription))); return
        }
        defer { _ = try? Self.detachDMG(mountPoint: mountPoint) }

        if canceled { finish(.failure(.canceled)); return }

        onPhase?(.copying)
        let appInDmg = (mountPoint as NSString).appendingPathComponent("LibreOffice.app")
        let destDir = NSHomeDirectory() + "/Library/Application Support/AnyView"
        let destApp = destDir + "/LibreOffice.app"
        let fm = FileManager.default
        try? fm.createDirectory(atPath: destDir, withIntermediateDirectories: true)
        try? fm.removeItem(atPath: destApp)

        let cp = Process()
        cp.executableURL = URL(fileURLWithPath: "/bin/cp")
        cp.arguments = ["-R", appInDmg, destApp]
        let cpErr = Pipe()
        cp.standardError = cpErr
        cp.standardOutput = FileHandle.nullDevice
        do {
            try cp.run()
            cp.waitUntilExit()
            if cp.terminationStatus != 0 {
                let msg = String(data: cpErr.fileHandleForReading.readDataToEndOfFile(),
                                 encoding: .utf8) ?? "exit \(cp.terminationStatus)"
                finish(.failure(.copyFailed(msg))); return
            }
        } catch {
            finish(.failure(.copyFailed(error.localizedDescription))); return
        }

        if canceled {
            try? fm.removeItem(atPath: destApp)
            finish(.failure(.canceled))
            return
        }

        onPhase?(.finalizing)
        // Clear quarantine xattr so Gatekeeper allows headless launch.
        let xattr = Process()
        xattr.executableURL = URL(fileURLWithPath: "/usr/bin/xattr")
        xattr.arguments = ["-dr", "com.apple.quarantine", destApp]
        xattr.standardOutput = FileHandle.nullDevice
        xattr.standardError = FileHandle.nullDevice
        try? xattr.run()
        xattr.waitUntilExit()

        let soffice = destApp + "/Contents/MacOS/soffice"
        guard fm.isExecutableFile(atPath: soffice) else {
            finish(.failure(.verificationFailed)); return
        }
        finish(.success(soffice))
    }

    private func finish(_ result: Result<String, InstallError>) {
        let cb = onFinish
        onFinish = nil
        onPhase = nil
        session?.invalidateAndCancel()
        session = nil
        task = nil
        cb?(result)
    }

    // MARK: URL + hdiutil

    static func dmgURL() -> URL? {
        let arch: String
        switch ProcessInfo.processInfo.machineArchitecture {
        case "arm64": arch = "aarch64"
        case "x86_64": arch = "x86_64"
        default: return nil
        }
        let v = version
        let str = "https://download.documentfoundation.org/libreoffice/stable/\(v)/mac/\(arch)/LibreOffice_\(v)_MacOS_\(arch).dmg"
        return URL(string: str)
    }

    static func attachDMG(at path: String) throws -> String {
        let p = Process()
        p.executableURL = URL(fileURLWithPath: "/usr/bin/hdiutil")
        p.arguments = ["attach", "-nobrowse", "-noautoopen", "-readonly", "-plist", path]
        let out = Pipe()
        let err = Pipe()
        p.standardOutput = out
        p.standardError = err
        try p.run()
        let outData = out.fileHandleForReading.readDataToEndOfFile()
        let errData = err.fileHandleForReading.readDataToEndOfFile()
        p.waitUntilExit()
        if p.terminationStatus != 0 {
            let msg = String(data: errData, encoding: .utf8) ?? "exit \(p.terminationStatus)"
            throw InstallError.mountFailed(msg)
        }
        guard let plist = try? PropertyListSerialization.propertyList(from: outData,
                                                                     options: [],
                                                                     format: nil) as? [String: Any],
              let entities = plist["system-entities"] as? [[String: Any]] else {
            throw InstallError.mountFailed("could not parse hdiutil output")
        }
        for entity in entities {
            if let mp = entity["mount-point"] as? String, !mp.isEmpty {
                return mp
            }
        }
        throw InstallError.mountFailed("no mount point in hdiutil output")
    }

    @discardableResult
    static func detachDMG(mountPoint: String) throws -> Bool {
        let p = Process()
        p.executableURL = URL(fileURLWithPath: "/usr/bin/hdiutil")
        p.arguments = ["detach", "-force", mountPoint]
        p.standardOutput = FileHandle.nullDevice
        p.standardError = FileHandle.nullDevice
        try p.run()
        p.waitUntilExit()
        return p.terminationStatus == 0
    }
}

private extension ProcessInfo {
    var machineArchitecture: String {
        var sysinfo = utsname()
        uname(&sysinfo)
        let mirror = Mirror(reflecting: sysinfo.machine)
        let chars = mirror.children.compactMap { ($0.value as? Int8) }
            .prefix(while: { $0 != 0 })
            .map { UInt8(bitPattern: $0) }
        return String(decoding: chars, as: UTF8.self)
    }
}

/// Tiny progress sheet. Programmatic — no XIB.
final class InstallProgressSheet {
    let window: NSWindow
    private let titleLabel: NSTextField
    private let detailLabel: NSTextField
    private let progress: NSProgressIndicator
    private let cancelButton: NSButton
    var onCancel: (() -> Void)?

    init() {
        let rect = NSRect(x: 0, y: 0, width: 460, height: 150)
        window = NSWindow(contentRect: rect,
                          styleMask: [.titled],
                          backing: .buffered,
                          defer: false)
        window.title = "安装 LibreOffice"

        titleLabel = NSTextField(labelWithString: "正在下载 LibreOffice…")
        titleLabel.font = NSFont.systemFont(ofSize: 13, weight: .semibold)
        titleLabel.translatesAutoresizingMaskIntoConstraints = false

        detailLabel = NSTextField(labelWithString: " ")
        detailLabel.font = NSFont.systemFont(ofSize: 11)
        detailLabel.textColor = .secondaryLabelColor
        detailLabel.translatesAutoresizingMaskIntoConstraints = false

        progress = NSProgressIndicator()
        progress.style = .bar
        progress.isIndeterminate = false
        progress.minValue = 0
        progress.maxValue = 1
        progress.translatesAutoresizingMaskIntoConstraints = false

        cancelButton = NSButton(title: "取消", target: nil, action: nil)
        cancelButton.bezelStyle = .rounded
        cancelButton.translatesAutoresizingMaskIntoConstraints = false

        let content = NSView(frame: rect)
        content.addSubview(titleLabel)
        content.addSubview(detailLabel)
        content.addSubview(progress)
        content.addSubview(cancelButton)
        window.contentView = content

        NSLayoutConstraint.activate([
            titleLabel.topAnchor.constraint(equalTo: content.topAnchor, constant: 20),
            titleLabel.leadingAnchor.constraint(equalTo: content.leadingAnchor, constant: 20),
            titleLabel.trailingAnchor.constraint(equalTo: content.trailingAnchor, constant: -20),

            progress.topAnchor.constraint(equalTo: titleLabel.bottomAnchor, constant: 14),
            progress.leadingAnchor.constraint(equalTo: content.leadingAnchor, constant: 20),
            progress.trailingAnchor.constraint(equalTo: content.trailingAnchor, constant: -20),

            detailLabel.topAnchor.constraint(equalTo: progress.bottomAnchor, constant: 8),
            detailLabel.leadingAnchor.constraint(equalTo: content.leadingAnchor, constant: 20),
            detailLabel.trailingAnchor.constraint(equalTo: content.trailingAnchor, constant: -20),

            cancelButton.bottomAnchor.constraint(equalTo: content.bottomAnchor, constant: -16),
            cancelButton.trailingAnchor.constraint(equalTo: content.trailingAnchor, constant: -20),
        ])

        cancelButton.target = self
        cancelButton.action = #selector(cancelTapped(_:))
    }

    func update(_ phase: LibreOfficeInstaller.Phase) {
        switch phase {
        case .downloading(let got, let total):
            titleLabel.stringValue = "正在下载 LibreOffice…"
            if total > 0 {
                progress.isIndeterminate = false
                progress.doubleValue = Double(got) / Double(total)
                detailLabel.stringValue = "\(formatMB(got)) / \(formatMB(total))"
            } else {
                progress.isIndeterminate = true
                progress.startAnimation(nil)
                detailLabel.stringValue = "\(formatMB(got))"
            }
        case .mounting:
            titleLabel.stringValue = "挂载 DMG…"
            detailLabel.stringValue = " "
            progress.isIndeterminate = true
            progress.startAnimation(nil)
        case .copying:
            titleLabel.stringValue = "复制 LibreOffice.app…"
            detailLabel.stringValue = "~/Library/Application Support/AnyView/"
            progress.isIndeterminate = true
            progress.startAnimation(nil)
        case .finalizing:
            titleLabel.stringValue = "清理隔离属性…"
            detailLabel.stringValue = " "
        }
    }

    @objc private func cancelTapped(_ sender: Any?) {
        cancelButton.isEnabled = false
        cancelButton.title = "取消中…"
        onCancel?()
    }

    private func formatMB(_ bytes: Int64) -> String {
        let mb = Double(bytes) / 1024.0 / 1024.0
        return String(format: "%.1f MB", mb)
    }
}
