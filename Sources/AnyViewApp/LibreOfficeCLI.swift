import CommonCrypto
import Foundation

/// Locates and invokes `soffice` (LibreOffice headless) for high-fidelity docx → PDF conversion.
enum LibreOfficeCLI {
    enum CLIError: LocalizedError {
        case notFound
        case executionFailed(String)
        case noOutput

        var errorDescription: String? {
            switch self {
            case .notFound:
                return "LibreOffice (soffice) not found. Install it to enable high-fidelity mode."
            case .executionFailed(let msg):
                return "LibreOffice conversion failed: \(msg)"
            case .noOutput:
                return "LibreOffice produced no PDF output."
            }
        }
    }

    /// Path where the in-app installer drops LibreOffice.
    static var managedAppPath: String {
        return NSHomeDirectory() + "/Library/Application Support/AnyView/LibreOffice.app"
    }

    /// Find the soffice binary path.
    /// Search order: app-managed install → /Applications → brew → PATH.
    static func findSoffice() -> String? {
        let candidates = [
            managedAppPath + "/Contents/MacOS/soffice",
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "/opt/homebrew/bin/soffice",
            "/usr/local/bin/soffice",
        ]
        for path in candidates {
            if FileManager.default.isExecutableFile(atPath: path) {
                return path
            }
        }
        return resolveViaWhich()
    }

    /// Convert `inputPath` to PDF in `outputDir`. Returns the PDF path.
    /// Caller is responsible for picking a stable, isolated `outputDir`
    /// (LibreOffice writes `<basename>.pdf` next to it).
    static func convertToPDF(inputPath: String, outputDir: String) throws -> String {
        guard let soffice = findSoffice() else { throw CLIError.notFound }

        let fm = FileManager.default
        try fm.createDirectory(atPath: outputDir, withIntermediateDirectories: true)

        let process = Process()
        process.executableURL = URL(fileURLWithPath: soffice)
        process.arguments = [
            "--headless",
            "--convert-to", "pdf",
            "--outdir", outputDir,
            inputPath,
        ]
        // Each conversion gets its own user profile so concurrent runs don't collide
        // and we don't inherit a stale user profile from a desktop LibreOffice session.
        let profileDir = NSTemporaryDirectory() + "AnyView-LO-" + UUID().uuidString
        try fm.createDirectory(atPath: profileDir, withIntermediateDirectories: true)
        process.environment = [
            "HOME": NSHomeDirectory(),
            "PATH": "/usr/bin:/bin",
            "USER_INSTALLATION": "file://\(profileDir)",
        ]

        let stdout = Pipe()
        let stderr = Pipe()
        process.standardOutput = stdout
        process.standardError = stderr

        try process.run()
        let outData = stdout.fileHandleForReading.readDataToEndOfFile()
        let errData = stderr.fileHandleForReading.readDataToEndOfFile()
        process.waitUntilExit()
        try? fm.removeItem(atPath: profileDir)

        if process.terminationStatus != 0 {
            let errString = String(data: errData, encoding: .utf8)
                ?? String(data: outData, encoding: .utf8)
                ?? "exit \(process.terminationStatus)"
            throw CLIError.executionFailed(errString)
        }

        let baseName = (inputPath as NSString).lastPathComponent
        let stem = (baseName as NSString).deletingPathExtension
        let pdfPath = (outputDir as NSString).appendingPathComponent(stem + ".pdf")
        guard fm.fileExists(atPath: pdfPath) else { throw CLIError.noOutput }
        return pdfPath
    }

    private static func resolveViaWhich() -> String? {
        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/usr/bin/which")
        process.arguments = ["soffice"]

        let pipe = Pipe()
        process.standardOutput = pipe
        process.standardError = FileHandle.nullDevice

        do {
            try process.run()
            let data = pipe.fileHandleForReading.readDataToEndOfFile()
            process.waitUntilExit()
            if process.terminationStatus == 0,
               let path = String(data: data, encoding: .utf8)?
                .trimmingCharacters(in: .whitespacesAndNewlines),
               !path.isEmpty,
               FileManager.default.isExecutableFile(atPath: path) {
                return path
            }
        } catch {}
        return nil
    }
}

/// Shared "convert this file to PDF for fidelity mode" helper used by every
/// renderer that supports fidelity. Wraps `LibreOfficeCLI.convertToPDF` with:
///  - .docmod / .doct → extract inner source.docx first
///  - cache lookup + write keyed on the outer file's mtime
enum LibreOfficeFidelity {
    /// Extensions that can be re-rendered through LibreOffice.
    /// xlsx/xls included so users can opt into "print preview" view via the
    /// fidelity toggle, even though the default xlsx path is SheetJS in WebRenderer.
    static let supportedExtensions: Set<String> = [
        "docx", "docmod", "doct",
        "pptx", "ppt",
        "xlsx", "xls",
    ]

    /// Run conversion (with cache) and return the PDF path. Throws `FidelityError`.
    static func preparePDF(for filePath: String) throws -> String {
        if let cached = FidelityCache.cachedPDFPath(for: filePath) {
            return cached
        }

        let ext = (filePath as NSString).pathExtension.lowercased()
        var inputForSoffice = filePath
        var tempExtractDir: String?

        if ext == "docmod" || ext == "doct" {
            let dir: String
            do {
                dir = try ZipExtractor.extract(zipPath: filePath)
            } catch {
                throw FidelityError.conversionFailed(error.localizedDescription)
            }
            tempExtractDir = dir
            let inner = dir + "/source.docx"
            guard FileManager.default.fileExists(atPath: inner) else {
                ZipExtractor.cleanup(tempDir: dir)
                throw FidelityError.noSourceDocx
            }
            inputForSoffice = inner
        }
        defer { if let d = tempExtractDir { ZipExtractor.cleanup(tempDir: d) } }

        let outputDir = NSTemporaryDirectory() + "AnyView-LO-out-" + UUID().uuidString
        do {
            let pdfPath = try LibreOfficeCLI.convertToPDF(inputPath: inputForSoffice,
                                                         outputDir: outputDir)
            let cached = FidelityCache.store(pdfPath: pdfPath, for: filePath) ?? pdfPath
            try? FileManager.default.removeItem(atPath: outputDir)
            return cached
        } catch let err as LibreOfficeCLI.CLIError {
            try? FileManager.default.removeItem(atPath: outputDir)
            switch err {
            case .notFound:
                throw FidelityError.sofficeNotFound
            case .executionFailed, .noOutput:
                throw FidelityError.conversionFailed(err.errorDescription ?? "unknown")
            }
        } catch {
            try? FileManager.default.removeItem(atPath: outputDir)
            throw FidelityError.conversionFailed(error.localizedDescription)
        }
    }
}

/// Caches converted PDFs by content hash + mtime so reopens are instant.
enum FidelityCache {
    static var cacheDir: String {
        return NSHomeDirectory() + "/Library/Caches/AnyView/fidelity"
    }

    /// Returns a cached pdf path for the given source file if available, else nil.
    /// `sourcePathForHash` is the file whose content/mtime we key on (for .docmod/.doct
    /// callers should pass the .docmod/.doct path itself, not the inner source.docx —
    /// the wrapper's mtime is what users see change).
    static func cachedPDFPath(for sourcePathForHash: String) -> String? {
        guard let key = cacheKey(for: sourcePathForHash) else { return nil }
        let candidate = (cacheDir as NSString).appendingPathComponent(key + ".pdf")
        return FileManager.default.fileExists(atPath: candidate) ? candidate : nil
    }

    /// Move a freshly-rendered pdf into the cache; returns the new cached path.
    static func store(pdfPath: String, for sourcePathForHash: String) -> String? {
        guard let key = cacheKey(for: sourcePathForHash) else { return nil }
        let fm = FileManager.default
        try? fm.createDirectory(atPath: cacheDir, withIntermediateDirectories: true)
        let dst = (cacheDir as NSString).appendingPathComponent(key + ".pdf")
        try? fm.removeItem(atPath: dst)
        do {
            try fm.moveItem(atPath: pdfPath, toPath: dst)
            return dst
        } catch {
            // Fall back to copy (cross-volume) then delete original
            do {
                try fm.copyItem(atPath: pdfPath, toPath: dst)
                try? fm.removeItem(atPath: pdfPath)
                return dst
            } catch {
                return nil
            }
        }
    }

    /// Cache key: 12-char path-sha + 12-char (size+mtime)-fingerprint.
    private static func cacheKey(for path: String) -> String? {
        let fm = FileManager.default
        guard let attrs = try? fm.attributesOfItem(atPath: path) else { return nil }
        let size = (attrs[.size] as? NSNumber)?.uint64Value ?? 0
        let mtime = (attrs[.modificationDate] as? Date)?.timeIntervalSince1970 ?? 0
        let pathHash = sha256Hex(path).prefix(12)
        let fingerprint = sha256Hex("\(size):\(mtime)").prefix(12)
        return "\(pathHash)-\(fingerprint)"
    }

    private static func sha256Hex(_ s: String) -> String {
        let data = Data(s.utf8)
        var hash = [UInt8](repeating: 0, count: 32)
        _ = data.withUnsafeBytes { buf in
            CC_SHA256(buf.baseAddress, CC_LONG(buf.count), &hash)
        }
        return hash.map { String(format: "%02x", $0) }.joined()
    }
}
