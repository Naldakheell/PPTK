import AppKit
import ApplicationServices
import Carbon
import CoreGraphics
import Foundation

/// PPTK keyboard helper. Single keys: T R U O C P G M N I and /.
/// Canvas sequences: A/S/D/H/V/., then a second key within one second.
/// H and V take + or -. Shift is kept only so that character can be typed.
/// / swaps shapes. Command+/ swaps text. . then T merges text boxes.
/// Command+Shift+C copies shape formatting. Command+Shift+V pastes it.
/// Option+. removes margins. Plain . still starts the merge prefix.
/// Command+] and Command+[ change font size, including while editing text.
/// Command+Shift+] brings forward. Command+Shift+[ sends backward.
/// The bracket keys are the physical ANSI keys, not the Arabic character.
/// \ saves the selected slides. ; opens the slide library.
/// [ emails the selected slides. That is the physical ANSI key, not the Arabic character.
/// \ and [ run only when slides are selected. ; runs on the canvas or thumbnails.
/// None of them run while editing text, and none of them use a modifier.
/// Command+[ and Command+Shift+[ stay on font size and send backward.
///
/// Event tap: session tap, inserted at the head, default (not listen-only)
/// so a matching bare key can be deleted. Every other event is returned unchanged.
/// Any uncertain gate returns the original event.
/// A focused text field, the Stock Images web UI, or a dialog/sheet that contains the focus
/// is passed through before that decision. Another open dialog does not block the slide.
/// Focus is read again on each key.
/// The text-versus-canvas decision is otherwise the selection type read at the moment of the key.

private let powerPointBundleID = "com.microsoft.Powerpoint"
private let macroName = "PPTK_DispatchShortcut"

/// ANSI letter keycodes. A shifted letter is not a shortcut.
/// "+" and "-" are translated from the active layout, because "+" is Shift on many layouts.
private let keyLetters: [Int64: String] = [
    0x00: "A",
    0x01: "S",
    0x02: "D",
    0x04: "H",
    0x05: "G",
    0x08: "C",
    0x09: "V",
    0x0B: "B",
    0x0D: "W",
    0x0F: "R",
    0x11: "T",
    0x1F: "O",
    0x20: "U",
    0x22: "I",
    0x23: "P",
    0x25: "L",
    0x2D: "N",
    0x2E: "M",
]

private let directTokens: Set<String> = ["T", "R", "U", "O", "C", "P", "G", "M", "N", "I", "/"]
private let prefixLetters: Set<String> = ["A", "S", "D", "H", "V", "."]
private let sequenceTokens: Set<String> = [
    "A L", "A R", "A T", "A B", "A M", "A C",
    "S W", "S H", "S S",
    "D H", "D V",
    "H +", "H -",
    "V +", "V -",
    ". T",
]
private let shiftKeycodes: Set<Int64> = [0x38, 0x3C]
private let escapeKeycode: Int64 = 0x35
private let prefixWindow: TimeInterval = 1.0

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

private struct KeyContext {
    var gate: FocusGate
    var show: ShowGate
    var surface: Surface
    var detail: String
}

private final class Helper: @unchecked Sendable {
    var tap: CFMachPort?
    private let queue = DispatchQueue(label: "sa.odel.pptk.keys.dispatch")
    private var repeatGate = FocusGate.unknown
    private var loggedAccessibility = false
    private var loggedAutomation = false
    private var prefix: String?
    private var prefixDeadline: Date?
    private var prefixTimer: DispatchWorkItem?
    private var swallowedKeycode: Int64?
    private var focusObserver: AXObserver?
    private var observedPID: pid_t = 0

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
        NSWorkspace.shared.notificationCenter.addObserver(
            forName: NSWorkspace.didActivateApplicationNotification,
            object: nil,
            queue: .main
        ) { [weak self] note in
            let app = note.userInfo?[NSWorkspace.applicationUserInfoKey] as? NSRunningApplication
            if app?.bundleIdentifier == powerPointBundleID {
                self?.installFocusWatch()
                return
            }
            self?.dropPrefix(timeout: false)
        }
        installFocusWatch()

        let source = CFMachPortCreateRunLoopSource(kCFAllocatorDefault, tap, 0)
        CFRunLoopAddSource(CFRunLoopGetCurrent(), source, .commonModes)
        CGEvent.tapEnable(tap: tap, enable: true)
        let timer = CFRunLoopTimerCreateWithHandler(
            kCFAllocatorDefault,
            CFAbsoluteTimeGetCurrent() + 0.25,
            0.25,
            0,
            0
        ) { [weak self] _ in
            self?.restoreTapIfNeeded()
        }
        if let timer {
            CFRunLoopAddTimer(CFRunLoopGetCurrent(), timer, .commonModes)
        }
        CFRunLoopRun()
        return 0
    }

    func restoreTapIfNeeded() {
        guard let tap, !CGEvent.tapIsEnabled(tap: tap) else { return }
        CGEvent.tapEnable(tap: tap, enable: true)
        fputs("[PPTKKeys] event tap re-enabled\n", stderr)
    }

    func noteFocusMoved() {
        if focusInsideDialog() || focusAcceptsTyping() {
            dropPrefix(timeout: false)
        }
    }

    private func installFocusWatch() {
        guard AXIsProcessTrusted() else { return }
        guard let app = NSRunningApplication.runningApplications(withBundleIdentifier: powerPointBundleID).first else { return }
        if observedPID == app.processIdentifier, focusObserver != nil { return }
        if let focusObserver {
            CFRunLoopRemoveSource(CFRunLoopGetCurrent(), AXObserverGetRunLoopSource(focusObserver), .commonModes)
            self.focusObserver = nil
        }
        var observer: AXObserver?
        let refcon = Unmanaged.passUnretained(self).toOpaque()
        guard AXObserverCreate(app.processIdentifier, focusObserverCallback, &observer) == .success,
              let observer else { return }
        let element = AXUIElementCreateApplication(app.processIdentifier)
        guard AXObserverAddNotification(
            observer,
            element,
            kAXFocusedUIElementChangedNotification as CFString,
            refcon
        ) == .success else { return }
        CFRunLoopAddSource(CFRunLoopGetCurrent(), AXObserverGetRunLoopSource(observer), .commonModes)
        focusObserver = observer
        observedPID = app.processIdentifier
    }

    /// True when this key is one PPTK might swallow. Unrelated keys skip the focus walk.
    private func ownsShortcut(_ event: CGEvent, keycode: Int64) -> Bool {
        if prefix != nil { return true }
        if commandSlashEvent(event) || formatChord(event) != nil || bracketChord(event) != nil
            || optionPeriodEvent(event) || bareSlideToken(event) != nil {
            return true
        }
        if plainPunctuation(event) != nil { return true }
        if !event.flags.intersection([.maskCommand, .maskShift, .maskAlternate, .maskControl]).isEmpty {
            return false
        }
        return keyLetters[keycode] != nil
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
        let isRepeat = event.getIntegerValueField(.keyboardEventAutorepeat) != 0
        let front = isPowerPointFrontmost()
        if front, ownsShortcut(event, keycode: keycode) {
            let inDialog = focusInsideDialog()
            if inDialog || focusAcceptsTyping() {
                dropPrefix(timeout: false)
                swallowedKeycode = nil
                repeatGate = .unknown
                if !isRepeat {
                    fputs(inDialog ? "[PPTKKeys] dialog, key passed\n" : "[PPTKKeys] text entry, key passed\n", stderr)
                }
                return Unmanaged.passUnretained(event)
            }
        }
        if isRepeat, keycode == swallowedKeycode, front {
            return nil
        }

        let letter = keyLetters[keycode]
        let modified = !event.flags.intersection([.maskCommand, .maskShift, .maskAlternate, .maskControl]).isEmpty
        let punct = plainPunctuation(event)

        if prefix != nil {
            if !front || prefixExpired() {
                dropPrefix(timeout: front && prefixExpired())
            } else if spacingShiftPassthrough(prefix: prefix ?? "", keycode: keycode, event: event) {
                return Unmanaged.passUnretained(event)
            } else if let symbol = spacingSecondKey(event), (prefix == "H" || prefix == "V") {
                let context = readContext(isRepeat: false, letter: symbol)
                let combined = "\(self.prefix ?? "") \(symbol)"
                if sequenceTokens.contains(combined), sequenceAllowed(context) {
                    dropPrefix(timeout: false)
                    fputs("[PPTKKeys] dispatch \(combined)\n", stderr)
                    queue.async { [weak self] in self?.dispatch(combined) }
                    swallowedKeycode = keycode
                    return nil
                }
                dropPrefix(timeout: false)
                swallowedKeycode = nil
                return Unmanaged.passUnretained(event)
            } else if commandSlashEvent(event) {
                dropPrefix(timeout: false)
                return handleCommandSlash(event: event, keycode: keycode, isRepeat: isRepeat)
            } else if let chord = formatChord(event) {
                dropPrefix(timeout: false)
                return handleCanvasToken(event: event, keycode: keycode, isRepeat: isRepeat, token: chord)
            } else if let token = bracketChord(event) {
                dropPrefix(timeout: false)
                return handleBracket(event: event, keycode: keycode, isRepeat: isRepeat, token: token)
            } else if optionPeriodEvent(event) {
                dropPrefix(timeout: false)
                return handleCanvasToken(event: event, keycode: keycode, isRepeat: isRepeat, token: "OPT .")
            } else if let token = bareSlideToken(event) {
                dropPrefix(timeout: false)
                return handleSlideToken(event: event, keycode: keycode, isRepeat: isRepeat, token: token)
            } else if let punct {
                let context = readContext(isRepeat: false, letter: punct)
                let combined = "\(self.prefix ?? "") \(punct)"
                if sequenceTokens.contains(combined), sequenceAllowed(context) {
                    dropPrefix(timeout: false)
                    fputs("[PPTKKeys] dispatch \(combined)\n", stderr)
                    queue.async { [weak self] in self?.dispatch(combined) }
                    swallowedKeycode = keycode
                    return nil
                }
                dropPrefix(timeout: false)
                return handleUnprefixed(token: punct, keycode: keycode, isRepeat: isRepeat, event: event, context: context)
            } else if modified || keycode == escapeKeycode || letter == nil {
                dropPrefix(timeout: false)
                swallowedKeycode = nil
                return Unmanaged.passUnretained(event)
            } else if let letter {
                let context = readContext(isRepeat: false, letter: letter)
                let combined = "\(self.prefix ?? "") \(letter)"
                if sequenceTokens.contains(combined), sequenceAllowed(context) {
                    dropPrefix(timeout: false)
                    fputs("[PPTKKeys] dispatch \(combined)\n", stderr)
                    queue.async { [weak self] in self?.dispatch(combined) }
                    swallowedKeycode = keycode
                    return nil
                }
                dropPrefix(timeout: false)
                return handleUnprefixed(token: letter, keycode: keycode, isRepeat: isRepeat, event: event, context: context)
            }
        }

        if commandSlashEvent(event) {
            guard front else {
                swallowedKeycode = nil
                return Unmanaged.passUnretained(event)
            }
            return handleCommandSlash(event: event, keycode: keycode, isRepeat: isRepeat)
        }

        if let chord = formatChord(event) {
            guard front else {
                swallowedKeycode = nil
                return Unmanaged.passUnretained(event)
            }
            return handleCanvasToken(event: event, keycode: keycode, isRepeat: isRepeat, token: chord)
        }

        if let token = bracketChord(event) {
            guard front else {
                swallowedKeycode = nil
                return Unmanaged.passUnretained(event)
            }
            return handleBracket(event: event, keycode: keycode, isRepeat: isRepeat, token: token)
        }

        if optionPeriodEvent(event) {
            guard front else {
                swallowedKeycode = nil
                return Unmanaged.passUnretained(event)
            }
            return handleCanvasToken(event: event, keycode: keycode, isRepeat: isRepeat, token: "OPT .")
        }

        if let token = bareSlideToken(event) {
            guard front else {
                swallowedKeycode = nil
                return Unmanaged.passUnretained(event)
            }
            return handleSlideToken(event: event, keycode: keycode, isRepeat: isRepeat, token: token)
        }

        if front, let punct {
            return handleUnprefixed(token: punct, keycode: keycode, isRepeat: isRepeat, event: event, context: nil)
        }

        guard front, !modified, let letter else {
            swallowedKeycode = nil
            return Unmanaged.passUnretained(event)
        }
        return handleUnprefixed(token: letter, keycode: keycode, isRepeat: isRepeat, event: event, context: nil)
    }

    private func handleCommandSlash(event: CGEvent, keycode: Int64, isRepeat: Bool) -> Unmanaged<CGEvent>? {
        let context = readContext(isRepeat: isRepeat, letter: "CMD /")
        guard directAllowed("CMD /", context) else {
            if context.gate == .unknown, !loggedAccessibility {
                loggedAccessibility = true
                fputs("Accessibility: \(context.detail). CMD / is not intercepted.\n", stderr)
            }
            swallowedKeycode = nil
            return Unmanaged.passUnretained(event)
        }
        if !isRepeat {
            fputs("[PPTKKeys] dispatch CMD /\n", stderr)
            queue.async { [weak self] in self?.dispatch("CMD /") }
        }
        swallowedKeycode = keycode
        return nil
    }

    /// Font size works on a shape or while editing text. Z-order uses the same gate:
    /// PowerPoint applies Bring Forward and Send Backward to the shape that contains the caret.
    private func handleBracket(event: CGEvent, keycode: Int64, isRepeat: Bool, token: String) -> Unmanaged<CGEvent>? {
        let context = readContext(isRepeat: isRepeat, letter: token)
        guard bracketAllowed(context) else {
            if context.gate == .unknown, !loggedAccessibility {
                loggedAccessibility = true
                fputs("Accessibility: \(context.detail). \(token) is not intercepted.\n", stderr)
            }
            swallowedKeycode = nil
            return Unmanaged.passUnretained(event)
        }
        if !isRepeat {
            fputs("[PPTKKeys] dispatch \(token)\n", stderr)
            queue.async { [weak self] in self?.dispatch(token) }
        }
        swallowedKeycode = keycode
        return nil
    }

    private func bracketAllowed(_ context: KeyContext) -> Bool {
        guard context.show == .editing else { return false }
        switch context.gate {
        case .canvas, .editing:
            break
        default:
            return false
        }
        if context.gate == .canvas, context.surface != .plain { return false }
        return !focusInsideExcludedPane()
    }

    /// \ and EMAIL SLIDES only when slides are selected, so neither shows "No slides selected."
    /// ; on the slide canvas or the thumbnail pane. A slide selection is not required.
    private func handleSlideToken(event: CGEvent, keycode: Int64, isRepeat: Bool, token: String) -> Unmanaged<CGEvent>? {
        let context = readContext(isRepeat: isRepeat, letter: token)
        let allowed: Bool
        switch token {
        case "\\", "EMAIL SLIDES":
            allowed = saveSlidesAllowed(context)
        default:
            allowed = libraryAllowed(context)
        }
        guard allowed else {
            if context.gate == .unknown, !loggedAccessibility {
                loggedAccessibility = true
                fputs("Accessibility: \(context.detail). \(token) is not intercepted.\n", stderr)
            }
            swallowedKeycode = nil
            return Unmanaged.passUnretained(event)
        }
        if !isRepeat {
            fputs("[PPTKKeys] dispatch \(token)\n", stderr)
            queue.async { [weak self] in self?.dispatch(token) }
        }
        swallowedKeycode = keycode
        return nil
    }

    private func saveSlidesAllowed(_ context: KeyContext) -> Bool {
        guard context.show == .editing, context.gate == .slides, context.surface != .blocked else { return false }
        return !focusInsideExcludedPane()
    }

    private func libraryAllowed(_ context: KeyContext) -> Bool {
        guard context.show == .editing else { return false }
        let onSurface: Bool
        switch (context.gate, context.surface) {
        case (.canvas, .plain), (.canvas, .thumbnail), (.slides, .plain), (.slides, .thumbnail):
            onSurface = true
        default:
            onSurface = false
        }
        return onSurface && !focusInsideExcludedPane()
    }

    private func handleCanvasToken(event: CGEvent, keycode: Int64, isRepeat: Bool, token: String) -> Unmanaged<CGEvent>? {
        let context = readContext(isRepeat: isRepeat, letter: token)
        guard directAllowed(token, context) else {
            if context.gate == .unknown, !loggedAccessibility {
                loggedAccessibility = true
                fputs("Accessibility: \(context.detail). \(token) is not intercepted.\n", stderr)
            }
            swallowedKeycode = nil
            return Unmanaged.passUnretained(event)
        }
        if !isRepeat {
            fputs("[PPTKKeys] dispatch \(token)\n", stderr)
            queue.async { [weak self] in self?.dispatch(token) }
        }
        swallowedKeycode = keycode
        return nil
    }

    private func handleUnprefixed(
        token: String,
        keycode: Int64,
        isRepeat: Bool,
        event: CGEvent,
        context supplied: KeyContext?
    ) -> Unmanaged<CGEvent>? {
        let isDirect = directTokens.contains(token)
        let isPrefix = prefixLetters.contains(token)
        guard isDirect || isPrefix else {
            swallowedKeycode = nil
            return Unmanaged.passUnretained(event)
        }
        let context = supplied ?? readContext(isRepeat: isRepeat, letter: token)
        if isPrefix {
            guard !isRepeat, sequenceAllowed(context) else {
                swallowedKeycode = isRepeat && sequenceAllowed(context) ? keycode : nil
                if isRepeat, sequenceAllowed(context) { return nil }
                return Unmanaged.passUnretained(event)
            }
            armPrefix(token)
            swallowedKeycode = keycode
            return nil
        }
        guard directAllowed(token, context) else {
            if context.gate == .unknown, !loggedAccessibility {
                loggedAccessibility = true
                fputs("Accessibility: \(context.detail). \(token) is not intercepted.\n", stderr)
            }
            swallowedKeycode = nil
            return Unmanaged.passUnretained(event)
        }
        if !isRepeat {
            if token == "T" {
                dispatch(token)
            } else {
                queue.async { [weak self] in self?.dispatch(token) }
            }
        }
        swallowedKeycode = keycode
        return nil
    }

    private func readContext(isRepeat: Bool, letter: String) -> KeyContext {
        if isRepeat, repeatGate != .unknown {
            return KeyContext(gate: repeatGate, show: .editing, surface: surfaceFor(repeatGate), detail: "repeat")
        }
        let read = modeRead()
        repeatGate = read.gate
        let context = KeyContext(gate: read.gate, show: read.show, surface: surfaceFor(read.gate), detail: read.detail)
        fputs("\(letter) \(context.gate) \(context.detail)\n", stderr)
        return context
    }

    private func surfaceFor(_ gate: FocusGate) -> Surface {
        switch gate {
        case .canvas, .slides:
            return surfaceVeto()
        case .editing, .notCanvas, .unknown:
            return .plain
        }
    }

    private func sequenceAllowed(_ context: KeyContext) -> Bool {
        context.gate == .canvas && context.surface == .plain && context.show == .editing && !focusInsideExcludedPane()
    }

    /// Ribbon and Notes often focus a child, not the pane itself. Single-key shortcuts keep the
    /// existing focused-element check. Sequences also refuse those panes.
    private func focusInsideExcludedPane() -> Bool {
        guard AXIsProcessTrusted(), let app = NSWorkspace.shared.frontmostApplication else { return false }
        let application = AXUIElementCreateApplication(app.processIdentifier)
        AXUIElementSetMessagingTimeout(application, 0.08)
        var error = AXError.success
        guard let raw = copyAttribute(application, kAXFocusedUIElementAttribute as String, error: &error),
              var element = axElement(raw) else { return false }
        var steps = 0
        while steps < 8 {
            AXUIElementSetMessagingTimeout(element, 0.08)
            let role = attribute(element, kAXRoleAttribute) ?? ""
            let subrole = attribute(element, kAXSubroleAttribute) ?? ""
            let description = (attribute(element, kAXDescriptionAttribute) ?? "").lowercased()
            if description.contains("notes pane") { return true }
            if role == "AXTabGroup", description.contains("ribbon") { return true }
            if role == "AXSheet" || role == "AXDialog" || subrole == "AXDialog" || subrole == "AXSystemDialog" || description.contains("alert") { return true }
            if role == "AXWindow" { return false }
            guard let next = parent(element) else { return false }
            element = next
            steps += 1
        }
        return false
    }

    private func directAllowed(_ token: String, _ context: KeyContext) -> Bool {
        let allowed = switch (context.gate, context.surface) {
        case (.canvas, .plain):
            true
        case (.slides, .thumbnail), (.slides, .plain):
            token == "N"
        default:
            false
        }
        return allowed && context.show == .editing
    }

    private func prefixExpired() -> Bool {
        guard let prefixDeadline else { return false }
        return Date() >= prefixDeadline
    }

    private func armPrefix(_ letter: String) {
        prefixTimer?.cancel()
        prefix = letter
        prefixDeadline = Date().addingTimeInterval(prefixWindow)
        fputs("[PPTKKeys] prefix \(letter)\n", stderr)
        let work = DispatchWorkItem { [weak self] in
            self?.dropPrefix(timeout: true)
        }
        prefixTimer = work
        DispatchQueue.main.asyncAfter(deadline: .now() + prefixWindow, execute: work)
    }

    private func dropPrefix(timeout: Bool) {
        guard prefix != nil else { return }
        prefixTimer?.cancel()
        prefixTimer = nil
        prefix = nil
        prefixDeadline = nil
        if timeout {
            fputs("[PPTKKeys] prefix timeout\n", stderr)
        }
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
        let parameter = appleScriptParameter(token)
        let source = """
        tell application "Microsoft PowerPoint"
            run VB macro macro name "\(macroName)" list of parameters {"\(parameter)"}
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

/// Shift is how many layouts type "+". The modifier key itself must not cancel H or V.
private func spacingShiftPassthrough(prefix: String, keycode: Int64, event: CGEvent) -> Bool {
    guard prefix == "H" || prefix == "V" else { return false }
    guard shiftKeycodes.contains(keycode) else { return false }
    return event.flags.intersection([.maskCommand, .maskAlternate, .maskControl]).isEmpty
}

/// "/" and "." come from the active layout, same as "+" and "-".
/// Command, Option, and Control are not part of the character. Shift is, when the layout needs it.
private func plainPunctuation(_ event: CGEvent) -> String? {
    if !event.flags.intersection([.maskCommand, .maskAlternate, .maskControl]).isEmpty {
        return nil
    }
    let keycode = UInt16(event.getIntegerValueField(.keyboardEventKeycode))
    let shift = event.flags.contains(.maskShift)
    guard let text = layoutCharacter(keycode: keycode, shift: shift) else { return nil }
    if text == "/" || text == "." { return text }
    return nil
}

/// Physical ANSI positions for \, ;, and [. Arabic layouts do not produce those characters here.
private let saveSlidesKeycode: Int64 = 0x2A
private let slideLibraryKeycode: Int64 = 0x29
private let emailSlidesKeycode: Int64 = 0x21

private func bareSlideToken(_ event: CGEvent) -> String? {
    if !event.flags.intersection([.maskCommand, .maskShift, .maskAlternate, .maskControl]).isEmpty {
        return nil
    }
    switch event.getIntegerValueField(.keyboardEventKeycode) {
    case saveSlidesKeycode:
        return "\\"
    case slideLibraryKeycode:
        return ";"
    case emailSlidesKeycode:
        return "EMAIL SLIDES"
    default:
        return nil
    }
}

/// A backslash would end the AppleScript string. Existing tokens have neither character.
private func appleScriptParameter(_ token: String) -> String {
    token.replacingOccurrences(of: "\\", with: "\\\\")
        .replacingOccurrences(of: "\"", with: "\\\"")
}

/// Physical ANSI bracket keys. Arabic layouts do not produce "[" or "]" on these keys.
/// Command alone changes font size. Command+Shift changes z-order by one step.
private let bracketTokens: [Int64: (plain: String, shifted: String)] = [
    0x1E: ("FONT +", "ORDER +"),
    0x21: ("FONT -", "ORDER -"),
]

private func bracketChord(_ event: CGEvent) -> String? {
    let flags = event.flags.intersection([.maskCommand, .maskShift, .maskAlternate, .maskControl])
    let keycode = event.getIntegerValueField(.keyboardEventKeycode)
    guard let token = bracketTokens[keycode] else { return nil }
    if flags == [.maskCommand] { return token.plain }
    if flags == [.maskCommand, .maskShift] { return token.shifted }
    return nil
}

/// Command+Shift+C and Command+Shift+V. Option and Control stay with macOS.
/// The letter is the same physical key as C and V, so Shift does not turn it into another character.
private func formatChord(_ event: CGEvent) -> String? {
    let flags = event.flags.intersection([.maskCommand, .maskShift, .maskAlternate, .maskControl])
    guard flags == [.maskCommand, .maskShift] else { return nil }
    let keycode = event.getIntegerValueField(.keyboardEventKeycode)
    guard let letter = keyLetters[keycode], letter == "C" || letter == "V" else { return nil }
    return "CMD SHIFT \(letter)"
}

/// Option+. only. The key is the one that types "." without Option.
private func optionPeriodEvent(_ event: CGEvent) -> Bool {
    guard event.flags.contains(.maskAlternate) else { return false }
    guard event.flags.intersection([.maskCommand, .maskControl, .maskShift]).isEmpty else { return false }
    let keycode = UInt16(event.getIntegerValueField(.keyboardEventKeycode))
    return layoutCharacter(keycode: keycode, shift: false) == "."
}

/// Command+/ only. Option and Control stay with macOS. Shift counts only when it is how "/" is typed.
private func commandSlashEvent(_ event: CGEvent) -> Bool {
    guard event.flags.contains(.maskCommand) else { return false }
    guard event.flags.intersection([.maskAlternate, .maskControl]).isEmpty else { return false }
    let keycode = UInt16(event.getIntegerValueField(.keyboardEventKeycode))
    let shift = event.flags.contains(.maskShift)
    return layoutCharacter(keycode: keycode, shift: shift) == "/"
}

/// The character this keydown inserts. Shift is part of the character. Command, Option, and Control are not.
/// UCKeyTranslate modifier bits are the EventRecord modifiers shifted right by 8, so Shift is 2.
private func spacingSecondKey(_ event: CGEvent) -> String? {
    if !event.flags.intersection([.maskCommand, .maskAlternate, .maskControl]).isEmpty {
        return nil
    }
    let keycode = UInt16(event.getIntegerValueField(.keyboardEventKeycode))
    let shift = event.flags.contains(.maskShift)
    guard let text = layoutCharacter(keycode: keycode, shift: shift) else { return nil }
    if text == "+" { return "+" }
    if text == "-" { return "-" }
    return nil
}

private func layoutCharacter(keycode: UInt16, shift: Bool) -> String? {
    guard let input = TISCopyCurrentKeyboardInputSource()?.takeRetainedValue() else { return nil }
    guard let raw = TISGetInputSourceProperty(input, kTISPropertyUnicodeKeyLayoutData) else { return nil }
    let data = Unmanaged<CFData>.fromOpaque(raw).takeUnretainedValue()
    guard let bytes = CFDataGetBytePtr(data) else { return nil }
    return bytes.withMemoryRebound(to: UCKeyboardLayout.self, capacity: 1) { layout in
        var deadKey: UInt32 = 0
        var length = 0
        var chars = [UniChar](repeating: 0, count: 4)
        let status = UCKeyTranslate(
            layout,
            keycode,
            UInt16(kUCKeyActionDown),
            shift ? 2 : 0,
            UInt32(LMGetKbdType()),
            OptionBits(kUCKeyTranslateNoDeadKeysBit),
            &deadKey,
            4,
            &length,
            &chars
        )
        guard status == noErr, length > 0 else { return nil }
        return String(utf16CodeUnits: chars, count: length)
    }
}

/// Focus sits inside a dialog or sheet. A dialog that is only open in the background does not count.
/// Paste Special and Open are AXWindow subrole AXDialog. Save As is an AXSheet on the document window.
/// A failed read returns false so a canvas shortcut is not blocked.
private func focusInsideDialog() -> Bool {
    guard AXIsProcessTrusted() else { return false }
    guard let app = NSWorkspace.shared.frontmostApplication,
          app.bundleIdentifier == powerPointBundleID else { return false }
    let application = AXUIElementCreateApplication(app.processIdentifier)
    AXUIElementSetMessagingTimeout(application, 0.05)
    var error = AXError.success
    guard let raw = copyAttribute(application, kAXFocusedUIElementAttribute as String, error: &error),
          var element = axElement(raw) else { return false }
    var steps = 0
    while steps < 12 {
        AXUIElementSetMessagingTimeout(element, 0.05)
        let role = attribute(element, kAXRoleAttribute) ?? ""
        let subrole = attribute(element, kAXSubroleAttribute) ?? ""
        if role == "AXSheet" || role == "AXDialog" || subrole == "AXDialog" || subrole == "AXSystemDialog" {
            return true
        }
        if role == "AXWindow" { return false }
        guard let next = parent(element) else { return false }
        element = next
        steps += 1
    }
    return false
}

/// Keyboard focus is inside a text control or the Stock Images web UI.
/// The slide editor returns false even while that pane is still open.
/// A failed read returns false so a canvas shortcut is not blocked.
private func focusAcceptsTyping() -> Bool {
    guard AXIsProcessTrusted() else { return false }
    guard let app = NSWorkspace.shared.frontmostApplication,
          app.bundleIdentifier == powerPointBundleID else { return false }
    let application = AXUIElementCreateApplication(app.processIdentifier)
    AXUIElementSetMessagingTimeout(application, 0.05)
    var error = AXError.success
    guard let raw = copyAttribute(application, kAXFocusedUIElementAttribute as String, error: &error),
          var element = axElement(raw) else { return false }
    var steps = 0
    while steps < 10 {
        AXUIElementSetMessagingTimeout(element, 0.05)
        let role = attribute(element, kAXRoleAttribute) ?? ""
        let subrole = attribute(element, kAXSubroleAttribute) ?? ""
        let description = (attribute(element, kAXDescriptionAttribute) ?? "").lowercased()
        let title = (attribute(element, kAXTitleAttribute) ?? "").lowercased()
        if textRoles.contains(role) || textRoles.contains(subrole) || subrole == "AXSearchField" {
            return true
        }
        if role == "AXWebArea" { return true }
        if description.contains("stock images") || title.contains("stock images") { return true }
        if role == "AXLayoutArea", description.hasPrefix(slideEditorPrefix.lowercased()) { return false }
        if role == "AXWindow" { return false }
        guard let next = parent(element) else { return false }
        element = next
        steps += 1
    }
    return false
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
    let result = helper.handle(proxy: proxy, type: type, event: event)
    helper.restoreTapIfNeeded()
    return result
}

private func focusObserverCallback(
    _ observer: AXObserver,
    _ element: AXUIElement,
    _ notification: CFString,
    _ refcon: UnsafeMutableRawPointer?
) {
    guard let refcon else { return }
    Unmanaged<Helper>.fromOpaque(refcon).takeUnretainedValue().noteFocusMoved()
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
