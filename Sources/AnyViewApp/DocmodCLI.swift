import Foundation

/// Locates and invokes the docmod CLI for rendering .docx files.
enum DocmodCLI {
    enum CLIError: LocalizedError {
        case notFound
        case executionFailed(String)

        var errorDescription: String? {
            switch self {
            case .notFound:
                return "docmod CLI not found. Install it to ~/.local/bin/docmod or add it to PATH."
            case .executionFailed(let msg):
                return "docmod render failed: \(msg)"
            }
        }
    }

    /// Find the docmod binary path.
    static func findDocmod() -> String? {
        // 1. Same directory as the .app bundle
        if let bundlePath = Bundle.main.executablePath {
            let bundleDir = (bundlePath as NSString).deletingLastPathComponent
            let candidate = (bundleDir as NSString).appendingPathComponent("docmod")
            if FileManager.default.isExecutableFile(atPath: candidate) {
                return candidate
            }
        }

        // 2. ~/.local/bin/docmod
        let localBin = NSHomeDirectory() + "/.local/bin/docmod"
        if FileManager.default.isExecutableFile(atPath: localBin) {
            return localBin
        }

        // 3. ~/.docmod/bin/docmod (legacy location)
        let homeBin = NSHomeDirectory() + "/.docmod/bin/docmod"
        if FileManager.default.isExecutableFile(atPath: homeBin) {
            return homeBin
        }

        // 4. DOCMOD_PATH environment variable
        if let envPath = ProcessInfo.processInfo.environment["DOCMOD_PATH"],
           FileManager.default.isExecutableFile(atPath: envPath) {
            return envPath
        }

        // 5. Check PATH via /usr/bin/which
        if let envPath = resolveViaEnv() {
            return envPath
        }

        // 6. Common install locations
        let commonPaths = [
            "/usr/local/bin/docmod",
            NSHomeDirectory() + "/.dotnet/tools/docmod",
        ]
        for path in commonPaths {
            if FileManager.default.isExecutableFile(atPath: path) {
                return path
            }
        }

        return nil
    }

    /// Run `docmod render <filePath>` and return the HTML output.
    static func render(filePath: String) throws -> String {
        guard let docmodPath = findDocmod() else {
            throw CLIError.notFound
        }

        let process = Process()
        process.executableURL = URL(fileURLWithPath: docmodPath)
        process.arguments = ["render", filePath]

        let stdout = Pipe()
        let stderr = Pipe()
        process.standardOutput = stdout
        process.standardError = stderr

        try process.run()

        // Read pipes before waitUntilExit to avoid deadlock when output exceeds pipe buffer
        let outData = stdout.fileHandleForReading.readDataToEndOfFile()
        let errData = stderr.fileHandleForReading.readDataToEndOfFile()
        process.waitUntilExit()

        if process.terminationStatus != 0 {
            let errString = String(data: errData, encoding: .utf8) ?? "Unknown error"
            throw CLIError.executionFailed(errString)
        }

        guard let html = String(data: outData, encoding: .utf8) else {
            throw CLIError.executionFailed("Could not decode output as UTF-8")
        }

        return html
    }

    /// Run `docmod create <output> --from <input>` to convert docx to docmod.
    static func createDocmod(from inputPath: String, to outputPath: String) throws {
        guard let docmodPath = findDocmod() else {
            throw CLIError.notFound
        }

        let process = Process()
        process.executableURL = URL(fileURLWithPath: docmodPath)
        process.arguments = ["create", outputPath, "--from", inputPath]

        let stderr = Pipe()
        process.standardOutput = FileHandle.nullDevice
        process.standardError = stderr

        try process.run()
        let errData = stderr.fileHandleForReading.readDataToEndOfFile()
        process.waitUntilExit()

        if process.terminationStatus != 0 {
            let errString = String(data: errData, encoding: .utf8) ?? "Unknown error"
            throw CLIError.executionFailed(errString)
        }
    }

    private static func resolveViaEnv() -> String? {
        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/usr/bin/which")
        process.arguments = ["docmod"]

        let pipe = Pipe()
        process.standardOutput = pipe
        process.standardError = FileHandle.nullDevice

        do {
            try process.run()
            let data = pipe.fileHandleForReading.readDataToEndOfFile()
            process.waitUntilExit()
            if process.terminationStatus == 0 {
                if let path = String(data: data, encoding: .utf8)?.trimmingCharacters(in: .whitespacesAndNewlines),
                   !path.isEmpty {
                    return path
                }
            }
        } catch {
            // Ignore
        }
        return nil
    }
}
