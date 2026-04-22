// swift-tools-version: 5.9
import PackageDescription

let package = Package(
    name: "AnyView",
    platforms: [.macOS(.v13)],
    products: [
        .executable(name: "AnyView", targets: ["AnyViewApp"]),
    ],
    targets: [
        .executableTarget(
            name: "AnyViewApp",
            path: "Sources/AnyViewApp",
            exclude: ["Info.plist"],
            resources: [.process("Resources")]
        ),
    ]
)
