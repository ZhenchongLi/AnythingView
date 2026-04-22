import Foundation

/// Extracts a zip file to a temporary directory.
enum ZipExtractor {
    /// Extract the zip file at `zipPath` to a new temporary directory.
    /// Returns the path to the temp directory.
    static func extract(zipPath: String) throws -> String {
        let tempDir = NSTemporaryDirectory() + "AnyView-" + UUID().uuidString
        try FileManager.default.createDirectory(atPath: tempDir, withIntermediateDirectories: true)

        let process = Process()
        process.executableURL = URL(fileURLWithPath: "/usr/bin/unzip")
        process.arguments = ["-o", "-q", zipPath, "-d", tempDir]
        process.standardOutput = FileHandle.nullDevice
        process.standardError = FileHandle.nullDevice

        try process.run()
        process.waitUntilExit()

        if process.terminationStatus != 0 {
            // Clean up on failure
            try? FileManager.default.removeItem(atPath: tempDir)
            throw NSError(domain: "ZipExtractor", code: 1,
                          userInfo: [NSLocalizedDescriptionKey: "Failed to extract zip file"])
        }

        return tempDir
    }

    /// Remove the temporary directory.
    static func cleanup(tempDir: String) {
        try? FileManager.default.removeItem(atPath: tempDir)
    }
}
