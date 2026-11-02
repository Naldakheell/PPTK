// swift-tools-version: 6.0
import PackageDescription

let package = Package(
    name: "PPTKKeys",
    platforms: [.macOS(.v13)],
    targets: [
        .executableTarget(
            name: "PPTKKeys",
            swiftSettings: [.swiftLanguageMode(.v5)],
            linkerSettings: [
                .linkedFramework("ApplicationServices"),
                .linkedFramework("AppKit")
            ]
        )
    ]
)
