import AppKit
import ApplicationServices
import CoreGraphics
import Foundation

/// PPTK keyboard helper for the first AUXI insert set: T R U O C P G M N.
///
/// Event tap: session tap, inserted at the head, default (not listen-only)
/// so a matching bare key can be deleted. Every other event is returned unchanged.
/// Any uncertain gate returns the original event.
/// The text-versus-canvas decision is the selection type read at the moment of the key.

private let powerPointBundleID = "com.microsoft.Powerpoint"
private let macroName = "PPTK_DispatchShortcut"

/// ANSI virtual keycodes. Shift is rejected separately, so these match the physical key.
private let shortcutTokens: [Int64: String] = [
    0x11: "T",
    0x0F: "R",
    0x20: "U",
    0x1F: "O",
    0x08: "C",
    0x23: "P",
    0x05: "G",
    0x2E: "M",
    0x2D: "N",
]

private let textRoles: Set<String> = [
    "AXTextArea", "AXTextField", "AXSearchField", "AXComboBox", "AXTextView",
]

/// Observed on Microsoft PowerPoint for Mac, normal view, English accessibility names:
/// canvas focus is AXLayoutArea "Slide Editor Pane, Slide N of M Slides".
/// Text editing moves focus to the parent AXScrollArea "scroll area".
/// Notes are AXLayoutArea "Notes Pane". Thumbnails are AXList "Thumbnail Pane".
/// The ribbon is AXTabGroup "ribbon". Slideshow focus is AXLayoutArea "layout area".
/// Only the slide-editor description is a shortcut. Anything else fails open.
private let slideEditorPrefix = "Slide Editor Pane"

private enum FocusGate {
    case canvas
    case slides
    case editing
    case notCanvas
    case unknown
}

private enum Surface {
    case plain
    case thumbnail
    case blocked
}

private enum ShowGate {
    case editing
    case slideshow
    case unknown
}

private final class Helper: @unchecked Sendable {
    var tap: CFMachPort?
    private let queue = DispatchQueue(label: "sa.odel.pptk.keys.dispatch")
    private var repeatGate = FocusGate.unknown
    private var loggedAccessibility = false
    private var loggedAutomation = false

    func start() -> Int32 {
        let trusted = AXIsProcessTrustedWithOptions(accessibilityPromptOptions)
        if trusted {
            fputs("Accessibility: granted\n", stdout)
        } else {
            fputs("Accessibility: not granted. R will type normally until it is granted in System Settings > Privacy & Security > Accessibility.\n", stderr)
        }

        let mask = CGEventMask(1 << CGEventType.keyDown.rawValue)
        let refcon = Unmanaged.passUnretained(self).toOpaque()
        guard let tap = CGEvent.tapCreate(
            tap: .cgSessionEventTap,
            place: .headInsertEventTap,
            options: .defaultTap,
            eventsOfInterest: mask,
            callback: tapCallback,
            userInfo: refcon
        ) else {
            fputs("Input Monitoring: not granted. No keys are intercepted. Enable it for this program in System Settings > Privacy & Security > Input Monitoring, then launch again.\n", stderr)
            fputs("Program path: \(CommandLine.arguments[0])\n", stderr)
            return 1
        }
        self.tap = tap
        fputs("Input Monitoring: event tap active (session, head, defaultTap, keyDown)\n", stdout)
        fputs("Mapped keys ask PowerPoint for the selection at the moment of the key.\n", stdout)
        fputs("Listening. Ctrl+C stops the helper.\n", stdout)

        let source = CFMachPortCreateRunLoopSource(kCFAllocatorDefault, tap, 0)
        CFRunLoopAddSource(CFRunLoopGetCurrent(), source, .commonModes)
        CGEvent.tapEnable(tap: tap, enable: true)
        CFRunLoopRun()
        return 0
    }

    func handle(proxy: CGEventTapProxy, type: CGEventType, event: CGEvent) -> Unmanaged<CGEvent>? {
        if type == .tapDisabledByTimeout || type == .tapDisabledByUserInput {
            if let tap {
                CGEvent.tapEnable(tap: tap, enable: true)
            }
            return Unmanaged.passUnretained(event)
        }
        guard type == .keyDown else { return Unmanaged.passUnretained(event) }
        let keycode = event.getIntegerValueField(.keyboardEventKeycode)
        guard let token = shortcutTokens[keycode] else {
            return Unmanaged.passUnretained(event)
        }
        let blocked = event.flags.intersection([.maskCommand, .maskShift, .maskAlternate, .maskControl])
        if !blocked.isEmpty { return Unmanaged.passUnretained(event) }
        guard isPowerPointFrontmost() else { return Unmanaged.passUnretained(event) }

        let isRepeat = event.getIntegerValueField(.keyboardEventAutorepeat) != 0
        let read: (gate: FocusGate, show: ShowGate, detail: String)
        if isRepeat, repeatGate != .unknown {
            read = (repeatGate, .editing, "repeat")
        } else {
            read = modeRead()
            repeatGate = read.gate
            fputs("\(token) \(read.gate) \(read.detail)\n", stderr)
        }

        let surface: Surface
        switch read.gate {
        case .canvas, .slides:
            surface = surfaceVeto()
        case .editing, .notCanvas, .unknown:
            surface = .plain
        }
        let allowed = switch (read.gate, surface) {
        case (.canvas, .plain):
            true
        case (.slides, .thumbnail), (.slides, .plain):
            token == "N"
        default:
            false
        }
        guard allowed, read.show == .editing else {
            if read.gate == .unknown, !loggedAccessibility {
                loggedAccessibility = true
                fputs("Accessibility: \(read.detail). \(token) is not intercepted.\n", stderr)
            }
            return Unmanaged.passUnretained(event)
        }

        if !isRepeat {
            if token == "T" {
                dispatch(token)
            } else {
                queue.async { [weak self] in self?.dispatch(token) }
            }
        }
        return nil
    }

    /// Asked on the key itself so the first R inside text is a letter, and the first R
    /// after leaving text inserts a shape. A background poll was always one step late.
    private func modeRead() -> (gate: FocusGate, show: ShowGate, detail: String) {
        switch appleScript(modeScript) {
        case .failed(let number, _):
            if number == -1743, !loggedAutomation {
                loggedAutomation = true
                fputs("Automation: not granted. The key is not intercepted. Allow this program to control Microsoft PowerPoint.\n", stderr)
            }
            return (.unknown, .unknown, "selection=unknown")
        case .value(let text) where text == "show":
            return (.notCanvas, .slideshow, "slideshow")
        case .value(let text) where text.contains("text"):
            return (.editing, .editing, "selection=text")
        case .value(let text) where text.contains("slide"):
            return (.slides, .editing, "selection=\(text)")
        case .value(let text) where text.contains("shape") || text.contains("none"):
            return (.canvas, .editing, "selection=\(text)")
        default:
            return (.unknown, .unknown, "selection=unknown")
        }
    }

    private func dispatch(_ token: String) {
        let source = """
        tell application "Microsoft PowerPoint"
            run VB macro macro name "\(macroName)" list of parameters {"\(token)"}
        end tell
        """
        switch appleScript(source) {
        case .value:
            fputs("Dispatched \(macroName) \(token)\n", stdout)
        case .failed(let number, let message):
            if number == -1743, !loggedAutomation {
                loggedAutomation = true
                fputs("Automation: not granted. The shortcut was not run.\n", stderr)
            } else {
                fputs("Dispatch failed \(number) \(message)\n", stderr)
            }
        }
    }
}

private func isPowerPointFrontmost() -> Bool {
    NSWorkspace.shared.frontmostApplication?.bundleIdentifier == powerPointBundleID
}

/// Positive identification only. A failed read does not block the canvas shortcut.
private func surfaceVeto() -> Surface {
    guard AXIsProcessTrusted() else { return .plain }
    guard let app = NSWorkspace.shared.frontmostApplication else { return .plain }
    let application = AXUIElementCreateApplication(app.processIdentifier)
    AXUIElementSetMessagingTimeout(application, 0.08)
    var error = AXError.success
    guard let raw = copyAttribute(application, kAXFocusedUIElementAttribute as String, error: &error),
          let focused = axElement(raw) else { return .plain }
    AXUIElementSetMessagingTimeout(focused, 0.08)
    let role = attribute(focused, kAXRoleAttribute) ?? ""
    let description = (attribute(focused, kAXDescriptionAttribute) ?? "").lowercased()
    if description.contains("thumbnail") { return .thumbnail }
    if description.contains("notes pane") { return .blocked }
    if role == "AXTabGroup", description.contains("ribbon") { return .blocked }
    if role == "AXSheet" || role == "AXDialog" || description.contains("alert") { return .blocked }
    return .plain
}

private func tapCallback(
    proxy: CGEventTapProxy,
    type: CGEventType,
    event: CGEvent,
    userInfo: UnsafeMutableRawPointer?
) -> Unmanaged<CGEvent>? {
    guard let userInfo else { return Unmanaged.passUnretained(event) }
    let helper = Unmanaged<Helper>.fromOpaque(userInfo).takeUnretainedValue()
    return helper.handle(proxy: proxy, type: type, event: event)
}

// MARK: - Accessibility

private struct FocusRead {
    var gate: FocusGate
    var detail: String
}

private func focusGate() -> FocusGate {
    readFocus().gate
}

private func readFocus() -> FocusRead {
    guard AXIsProcessTrusted() else {
        return FocusRead(gate: .unknown, detail: "Accessibility not trusted")
    }
    var axError = AXError.success
    guard let focused = focusedElement(error: &axError) else {
        return FocusRead(gate: .unknown, detail: "AX error \(axError.rawValue)")
    }
    let role = attribute(focused, kAXRoleAttribute) ?? ""
    let description = attribute(focused, kAXDescriptionAttribute) ?? ""
    let detail = "role=\(role) desc=\(description)"
    if textRoles.contains(role) {
        return FocusRead(gate: .editing, detail: detail)
    }
    if isThumbnail(role, description) {
        return thumbnailGate(detail)
    }
    if isExcludedPane(role, description) {
        return FocusRead(gate: .notCanvas, detail: detail)
    }
    if role == "AXLayoutArea", description.hasPrefix(slideEditorPrefix) {
        return FocusRead(gate: .canvas, detail: detail)
    }
    if let parentGate = ancestorGate(of: focused, detail: detail) {
        return parentGate
    }
    // After a shape insert, focus leaves the layout area and sits on its scroll area.
    // That same scroll area is also text editing, so the selection type breaks the tie.
    if role == "AXScrollArea" || description.isEmpty {
        if treeHasSlideEditor(focused) {
            switch selectionKind() {
            case .text:
                return FocusRead(gate: .editing, detail: detail + " selection=text")
            case .shapes, .none, .slides:
                return FocusRead(gate: .canvas, detail: detail + " selection=object")
            case .unknown:
                return FocusRead(gate: .unknown, detail: detail + " selection=unknown")
            }
        }
    }
    return FocusRead(gate: .notCanvas, detail: detail)
}

private func ancestorGate(of focused: AXUIElement, detail: String) -> FocusRead? {
    var current = parent(focused)
    var steps = 0
    while let element = current, steps < 6 {
        let role = attribute(element, kAXRoleAttribute) ?? ""
        let description = attribute(element, kAXDescriptionAttribute) ?? ""
        if textRoles.contains(role) {
            return FocusRead(gate: .editing, detail: detail)
        }
        if isThumbnail(role, description) {
            return thumbnailGate(detail)
        }
        if isExcludedPane(role, description) {
            return FocusRead(gate: .notCanvas, detail: detail)
        }
        if role == "AXLayoutArea", description.hasPrefix(slideEditorPrefix) {
            return FocusRead(gate: .canvas, detail: detail)
        }
        if role == "AXWindow" { return nil }
        current = parent(element)
        steps += 1
    }
    return nil
}

private func isThumbnail(_ role: String, _ description: String) -> Bool {
    let label = description.lowercased()
    return (role == "AXList" || role == "AXOutline") && label.contains("thumbnail")
}

/// Delete moves keyboard focus onto the thumbnail list even though the slide is still the working surface.
/// A selected slide thumbnail uses the same list, so R stays available there unless text is being edited.
private func thumbnailGate(_ detail: String) -> FocusRead {
    switch selectionKind() {
    case .text:
        return FocusRead(gate: .editing, detail: detail + " selection=text")
    case .unknown:
        return FocusRead(gate: .notCanvas, detail: detail + " selection=unknown")
    case .shapes, .none, .slides:
        return FocusRead(gate: .canvas, detail: detail + " selection=canvas")
    }
}

private func isExcludedPane(_ role: String, _ description: String) -> Bool {
    let label = description.lowercased()
    if label.contains("notes pane") { return true }
    if role == "AXTabGroup", label.contains("ribbon") { return true }
    return false
}

private func treeHasSlideEditor(_ root: AXUIElement) -> Bool {
    for child in children(root).prefix(8) {
        let role = attribute(child, kAXRoleAttribute) ?? ""
        let description = attribute(child, kAXDescriptionAttribute) ?? ""
        if role == "AXLayoutArea", description.hasPrefix(slideEditorPrefix) { return true }
        if isExcludedPane(role, description) { return false }
        if role == "AXScrollArea" || role == "AXGroup" {
            for grand in children(child).prefix(4) {
                let grandRole = attribute(grand, kAXRoleAttribute) ?? ""
                let grandDescription = attribute(grand, kAXDescriptionAttribute) ?? ""
                if grandRole == "AXLayoutArea", grandDescription.hasPrefix(slideEditorPrefix) { return true }
                if isExcludedPane(grandRole, grandDescription) { return false }
            }
        }
    }
    return false
}

private enum SelectionKind {
    case text
    case shapes
    case none
    case slides
    case unknown
}

private func selectionKind() -> SelectionKind {
    switch appleScript(modeScript) {
    case .value(let text) where text.contains("text"):
        return .text
    case .value(let text) where text.contains("slide"):
        return .slides
    case .value(let text) where text.contains("shape"):
        return .shapes
    case .value(let text) where text.contains("none"):
        return .none
    default:
        return .unknown
    }
}

private func focusedElement() -> AXUIElement? {
    var error = AXError.success
    return focusedElement(error: &error)
}

/// Prefer the frontmost app. System-wide focus fails for some apps even when that app answers.
private func focusedElement(error: inout AXError) -> AXUIElement? {
    if let app = NSWorkspace.shared.frontmostApplication {
        let element = AXUIElementCreateApplication(app.processIdentifier)
        AXUIElementSetMessagingTimeout(element, 0.4)
        if let raw = copyAttribute(element, kAXFocusedUIElementAttribute as String, error: &error),
           let focused = axElement(raw) {
            return focused
        }
    }
    let system = AXUIElementCreateSystemWide()
    AXUIElementSetMessagingTimeout(system, 0.4)
    guard let raw = copyAttribute(system, kAXFocusedUIElementAttribute as String, error: &error) else {
        return nil
    }
    return axElement(raw)
}

private func parent(_ element: AXUIElement) -> AXUIElement? {
    guard let raw = copyAttribute(element, kAXParentAttribute as String) else { return nil }
    return axElement(raw)
}

private func ancestor(of element: AXUIElement, role: String) -> AXUIElement? {
    var current: AXUIElement? = element
    var steps = 0
    while let item = current, steps < 24 {
        if attribute(item, kAXRoleAttribute) == role { return item }
        current = parent(item)
        steps += 1
    }
    return nil
}

private func elementArea(_ element: AXUIElement) -> CGFloat? {
    guard let raw = copyAttribute(element, kAXSizeAttribute as String) else { return nil }
    guard CFGetTypeID(raw) == AXValueGetTypeID() else { return nil }
    var size = CGSize.zero
    let value = raw as! AXValue
    guard AXValueGetValue(value, .cgSize, &size) else { return nil }
    let area = size.width * size.height
    return area > 0 ? area : nil
}

private func children(_ element: AXUIElement) -> [AXUIElement] {
    guard let raw = copyAttribute(element, kAXChildrenAttribute as String) else { return [] }
    guard CFGetTypeID(raw) == CFArrayGetTypeID() else { return [] }
    let array = raw as! CFArray
    var result: [AXUIElement] = []
    for index in 0..<CFArrayGetCount(array) {
        guard let pointer = CFArrayGetValueAtIndex(array, index) else { continue }
        let element = Unmanaged<AXUIElement>.fromOpaque(pointer).takeUnretainedValue()
        result.append(element)
    }
    return result
}

private func attribute(_ element: AXUIElement, _ name: String) -> String? {
    guard let raw = copyAttribute(element, name) else { return nil }
    return raw as? String
}

private func copyAttribute(_ element: AXUIElement, _ name: String) -> CFTypeRef? {
    var error = AXError.success
    return copyAttribute(element, name, error: &error)
}

private func copyAttribute(_ element: AXUIElement, _ name: String, error: inout AXError) -> CFTypeRef? {
    var value: CFTypeRef?
    error = AXUIElementCopyAttributeValue(element, name as CFString, &value)
    guard error == .success else { return nil }
    return value
}

private func axElement(_ raw: CFTypeRef) -> AXUIElement? {
    guard CFGetTypeID(raw) == AXUIElementGetTypeID() else { return nil }
    return (raw as! AXUIElement)
}

// MARK: - Apple events

private let modeScript = """
tell application "Microsoft PowerPoint"
    if (count of slide show windows) > 0 then
        return "show"
    end if
    try
        return selection type of selection of active window as text
    on error
        return "unknown"
    end try
end tell
"""

private enum ScriptResult {
    case value(String)
    case failed(Int, String)
}

private let scriptLock = NSLock()

private func appleScript(_ source: String) -> ScriptResult {
    scriptLock.lock()
    defer { scriptLock.unlock() }
    guard let script = NSAppleScript(source: source) else {
        return .failed(-1, "script was not created")
    }
    var error: NSDictionary?
    let descriptor = script.executeAndReturnError(&error)
    if let error {
        let number = error[NSAppleScript.errorNumber] as? Int ?? -1
        let message = error[NSAppleScript.errorMessage] as? String ?? ""
        return .failed(number, message)
    }
    return .value(descriptor.stringValue ?? "")
}

// MARK: - Describe

private let accessibilityPromptOptions = ["AXTrustedCheckOptionPrompt": true] as CFDictionary

private func describeFocus() -> Int32 {
    guard AXIsProcessTrustedWithOptions(accessibilityPromptOptions) else {
        fputs("Accessibility: not granted. Cannot read the focused element.\n", stderr)
        return 1
    }
    guard let focused = focusedElement() else {
        fputs("No focused element.\n", stderr)
        return 1
    }
    fputs("frontmost=\(NSWorkspace.shared.frontmostApplication?.bundleIdentifier ?? "none")\n", stdout)
    fputs("gate=\(String(describing: focusGate()))\n", stdout)
    fputs("focused:\n", stdout)
    printElement(focused, indent: "  ")
    fputs("ancestors:\n", stdout)
    var current = parent(focused)
    var steps = 0
    while let element = current, steps < 16 {
        printElement(element, indent: "  ")
        if attribute(element, kAXRoleAttribute) == "AXWindow" { break }
        current = parent(element)
        steps += 1
    }
    if let window = ancestor(of: focused, role: "AXWindow") {
        fputs("scroll areas in window:\n", stdout)
        describeScrollAreas(in: window, limit: 400)
    }
    return 0
}

private func printElement(_ element: AXUIElement, indent: String) {
    let fields = [
        "role": kAXRoleAttribute as String,
        "subrole": kAXSubroleAttribute as String,
        "roleDescription": kAXRoleDescriptionAttribute as String,
        "identifier": kAXIdentifierAttribute as String,
        "description": kAXDescriptionAttribute as String,
        "title": kAXTitleAttribute as String,
        "help": kAXHelpAttribute as String,
    ]
    var parts: [String] = []
    for (label, name) in fields {
        if let value = attribute(element, name), !value.isEmpty {
            parts.append("\(label)=\(value)")
        }
    }
    if let size = elementArea(element) {
        parts.append(String(format: "area=%.0f", size))
    }
    fputs(indent + parts.joined(separator: " ") + "\n", stdout)
}

private func describeScrollAreas(in window: AXUIElement, limit: Int) {
    var queue: [(AXUIElement, Int)] = [(window, 0)]
    var seen = 0
    while !queue.isEmpty, seen < limit {
        let (element, depth) = queue.removeFirst()
        seen += 1
        if attribute(element, kAXRoleAttribute) == "AXScrollArea" {
            printElement(element, indent: "  ")
        }
        if depth >= 8 { continue }
        for child in children(element) {
            queue.append((child, depth + 1))
        }
    }
}

// MARK: - Entry

private let helper = Helper()

switch CommandLine.arguments.dropFirst().first {
case "--describe-focus":
    exit(describeFocus())
case "--help":
    fputs("Usage: PPTKKeys [--describe-focus]\nCtrl+C stops the listener.\n", stdout)
    exit(0)
default:
    exit(helper.start())
}
