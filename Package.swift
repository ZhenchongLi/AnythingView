// swift-tools-version: 5.9
import PackageDescription

let package = Package(
    name: "ViewAnything",
    platforms: [.macOS(.v13)],
    targets: [
        .executableTarget(
            name: "ViewAnything",
            path: "Sources/ViewAnything",
            exclude: ["Info.plist"],
            resources: [.process("Resources")]
        ),
    ]
)
