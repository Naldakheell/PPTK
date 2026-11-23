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
    0x06: "Z",
    0x07: "X",
    0x08: "C",
    0x09: "V",
    0x0B: "B",
    0x0C: "Q",
    0x0E: "E",
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

private let directTokens: Set<String> = ["T", "R", "U", "O", "C", "P", "G", "M", "N", "I", "/", "E", "Q"]
private let prefixLetters: Set<String> = ["A", "S", "D", "H", "V", ".", "X"]
private let sequenceTokens: Set<String> = [
    "A L", "A R", "A T", "A B", "A M", "A C", "A A", "A P", "A X", "A Z",
    "S W", "S H", "S S",
    "D H", "D V",
    "H +", "H -",
    "V +", "V -",
    ". T",
    "X T",
]
private let shiftKeycodes: Set<Int64> = [0x38, 0x3C]
/// Physical top-row digits 2-7 (ANSI key codes).
/// Keypad 2-7 count too (Windows ShortcutChords accepts the numpad digits).
private let digitKeycodes: [Int64: Int] = [
    0x13: 2, 0x14: 3, 0x15: 4, 0x17: 5, 0x16: 6, 0x1A: 7,
    0x54: 2, 0x55: 3, 0x56: 4, 0x57: 5, 0x58: 6, 0x59: 7,
]
/// US period and slash positions. When the layout types a letter there, they keep the US meaning (Windows KeyLayout.KeepUsKey).
private let usPeriodKeycode: UInt16 = 0x2F
private let usSlashKeycode: UInt16 = 0x2C
/// Windows ignores injected keys. Tests may opt in to synthetic keys with this environment variable.
private let allowSyntheticKeys = ProcessInfo.processInfo.environment["PPTK_ALLOW_SYNTHETIC_KEYS"] == "1"
private let lockPath = NSHomeDirectory() + "/Library/Application Support/PPTK/PPTKKeys.lock"
private let gridTokens: Set<String> = ["CTRL 2", "CTRL 3", "OPT 2", "OPT 3"]
private let gridPacingMicroseconds: useconds_t = 100_000
private let escapeKeycode: Int64 = 0x35
private let prefixWindow: TimeInterval = 1.0

private typealias ModeRead = (gate: FocusGate, show: ShowGate, detail: String, pane: PaneKind)
/// Marks keys this helper replays so the tap lets them through. ASCII "PPTK".
private let replayMarker: Int64 = 0x5050_544B
/// Same cap as the Windows helper's QueueLimit.
private let heldKeyLimit = 128
/// A held key falls back to typing when PowerPoint does not answer in time.
private let readQueueLimit: TimeInterval = 3.0
private let readLimit: TimeInterval = 1.0
private let textboxLimit: TimeInterval = 3.0
private let replayInFlightLimit: TimeInterval = 1.0

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

/// PowerPoint's active pane, read with the selection (Windows ShortcutGate.Classify uses the same view types).
private enum PaneKind {
    case slide
    case thumbnails
    case sorter
    case blocked
    case unknown
}

private struct KeyContext {
    var gate: FocusGate
    var show: ShowGate
    var surface: Surface
    var detail: String
    var pane: PaneKind = .unknown
}

private final class Helper: @unchecked Sendable {
    var tap: CFMachPort?
    private let queue = DispatchQueue(label: "sa.odel.pptk.keys.dispatch")
    private let deploymentQueue = DispatchQueue(label: "sa.odel.pptk.keys.deployment")
    private var repeatGate = FocusGate.unknown
    private var loggedAccessibility = false
    private var loggedAutomation = false
    private var loggedReadFailure = false
    private var prefix: String?
    private var prefixDeadline: Date?
    private var prefixTimer: DispatchWorkItem?
    private var swallowedKeycode: Int64?
    private var focusObserver: AXObserver?
    private var observedPID: pid_t = 0

    // Held keys. The tap never waits for PowerPoint. A key that needs the selection is held,
    // the selection is read on `queue`, and held keys are then handled or replayed in order.
    // Same idea as the Windows helper's text capture after T (_capture / _pending / ReplayGuard).
    private var pending: [CGEvent] = []
    private var waiting = false
    private var replayVerbatim = false
    private var isHandling = false
    private var suppliedRead: ModeRead?
    private var generation = 0
    private var waitTimer: DispatchWorkItem?
    private var replayVerbatimPending = false
    // Replayed keys not yet seen back at the tap. A live key must not overtake them.
    private var replaysInFlight = 0
    private var lastReplayAt = Date.distantPast
    private var repeatPane = PaneKind.unknown
    // T replay target (Windows ReplayGuard): the PowerPoint process and window that were active at T.
    private var replayOriginPID: pid_t = 0
    private var replayOriginWindow: AXUIElement?
    private var lockDescriptor: Int32 = -1
    private var sourceLogCount = 0

    func start() -> Int32 {
        guard acquireSingleInstanceLock() else {
            fputs("PPTKKeys is already running. Only one helper may run at a time.\n", stderr)
            return 2
        }
        if allowSyntheticKeys {
            fputs("Synthetic keys: ACCEPTED (PPTK_ALLOW_SYNTHETIC_KEYS=1, test mode only)\n", stderr)
        } else {
            fputs("Synthetic keys: ignored (keys posted by other programs pass through)\n", stdout)
        }
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
        deploymentCheck()

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

    /// Packaged installs only (PPTK.pkg): the add-in itself is installed by the package into Office's system-wide
    /// Startup folder, where PowerPoint loads it for every user. The helper only copies the Email Slides AppleScript
    /// plugin into this user's PowerPoint scripts folder. Does nothing for a development install.
    private func deploymentCheck() {
        deploymentQueue.async {
            let result = syncPowerPointPlugin()
            if result != "not deployed" { fputs("[PPTKKeys] email plugin: \(result)\n", stdout) }
        }
    }

    /// One helper per user session (Windows uses a named mutex). The kernel drops the lock when the process ends.
    private func acquireSingleInstanceLock() -> Bool {
        let folder = (lockPath as NSString).deletingLastPathComponent
        try? FileManager.default.createDirectory(atPath: folder, withIntermediateDirectories: true)
        let fd = open(lockPath, O_CREAT | O_RDWR, 0o644)
        guard fd >= 0 else { return true }
        if flock(fd, LOCK_EX | LOCK_NB) != 0 {
            close(fd)
            return false
        }
        lockDescriptor = fd
        return true
    }

    /// Posted by another program (not the keyboard, not this helper). Windows: LLKHF_INJECTED.
    private func isSynthetic(_ event: CGEvent) -> Bool {
        guard !allowSyntheticKeys else { return false }
        let pid = event.getIntegerValueField(.eventSourceUnixProcessID)
        return pid != 0 && pid != Int64(getpid())
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
            || optionPeriodEvent(event) || bareSlideToken(event) != nil || digitChord(event) != nil {
            return true
        }
        if plainPunctuation(event) != nil { return true }
        if !event.flags.intersection([.maskCommand, .maskShift, .maskAlternate, .maskControl]).isEmpty {
            return false
        }
        return keyLetters[keycode] != nil
    }

    /// Tap entry. Returns at once: no AppleScript runs here, so a nested callback cannot deadlock.
    func intercept(type: CGEventType, event: CGEvent) -> Unmanaged<CGEvent>? {
        guard type == .keyDown else { return handle(type: type, event: event) }
        if event.getIntegerValueField(.eventSourceUserData) == replayMarker {
            if replaysInFlight > 0 { replaysInFlight -= 1 }
            return Unmanaged.passUnretained(event)
        }
        if sourceLogCount < 5, isPowerPointFrontmost() {
            sourceLogCount += 1
            fputs("[PPTKKeys] key source pid=\(event.getIntegerValueField(.eventSourceUnixProcessID)) state=\(event.getIntegerValueField(.eventSourceStateID))\n", stderr)
        }
        defer { scheduleDrainIfIdle() }
        if isHandling || waiting || !pending.isEmpty || replayStillInFlight() {
            // Held in order; a synthetic key is replayed as typed, never read as a shortcut (see drain).
            return hold(event)
        }
        if isSynthetic(event) {
            return Unmanaged.passUnretained(event)
        }
        if needsSelectionRead(event) {
            guard let copy = event.copy() else { return Unmanaged.passUnretained(event) }
            pending.append(copy)
            startRead()
            return nil
        }
        isHandling = true
        defer { isHandling = false }
        return handle(type: type, event: event)
    }

    /// Keeps arrival order while an earlier key is still being decided.
    private func hold(_ event: CGEvent) -> Unmanaged<CGEvent>? {
        guard isPowerPointFrontmost() else {
            discardPending("app changed")
            return Unmanaged.passUnretained(event)
        }
        guard pending.count < heldKeyLimit, let copy = event.copy() else {
            fputs("[PPTKKeys] hold limit, key passed\n", stderr)
            return Unmanaged.passUnretained(event)
        }
        pending.append(copy)
        return nil
    }

    /// Same keys that used to ask PowerPoint inside the tap. Text entry, dialogs,
    /// and cached auto-repeat are decided without AppleScript and stay immediate.
    private func needsSelectionRead(_ event: CGEvent) -> Bool {
        guard isPowerPointFrontmost() else { return false }
        let keycode = event.getIntegerValueField(.keyboardEventKeycode)
        guard ownsShortcut(event, keycode: keycode) else { return false }
        if event.getIntegerValueField(.keyboardEventAutorepeat) != 0,
           keycode == swallowedKeycode || repeatGate != .unknown {
            return false
        }
        return !(focusInsideDialog() || focusAcceptsTyping())
    }

    private func startRead() {
        waiting = true
        generation += 1
        let token = generation
        armWaitTimer(token, after: readQueueLimit)
        queue.async { [weak self] in
            DispatchQueue.main.async { self?.readStarted(token) }
            guard let read = self?.modeRead() else { return }
            DispatchQueue.main.async { self?.readArrived(token, read) }
        }
    }

    /// The read waits behind earlier macros on `queue`. Its own limit starts when it runs.
    private func readStarted(_ token: Int) {
        guard token == generation, waiting else { return }
        armWaitTimer(token, after: readLimit)
    }

    private func readArrived(_ token: Int, _ read: ModeRead) {
        guard token == generation, waiting else { return }
        resolveFirst(read)
    }

    private func armWaitTimer(_ token: Int, after seconds: TimeInterval) {
        waitTimer?.cancel()
        let work = DispatchWorkItem { [weak self] in self?.waitTimedOut(token) }
        waitTimer = work
        DispatchQueue.main.asyncAfter(deadline: .now() + seconds, execute: work)
    }

    private func waitTimedOut(_ token: Int) {
        guard token == generation, waiting else { return }
        if replayVerbatimPending {
            // T macro did not finish. Replay only back into PowerPoint, never into a dialog (Windows ReplayGuard).
            replayVerbatimPending = false
            endWait()
            if replayTargetValid() {
                fputs("[PPTKKeys] T slow, replaying held keys\n", stderr)
                replayVerbatim = true
                drain()
            } else {
                discardPending("T timeout")
            }
            return
        }
        fputs("[PPTKKeys] selection read timed out, key typed\n", stderr)
        resolveFirst((.unknown, .unknown, "selection=timeout", .unknown))
    }

    private func endWait() {
        generation += 1
        waitTimer?.cancel()
        waitTimer = nil
        waiting = false
    }

    private func resolveFirst(_ read: ModeRead) {
        endWait()
        guard !pending.isEmpty else { return }
        guard isPowerPointFrontmost() else {
            discardPending("app changed")
            return
        }
        process(pending.removeFirst(), read: read)
        drain()
    }

    /// Runs the unchanged gate with a selection read that already happened.
    private func process(_ event: CGEvent, read: ModeRead?) {
        suppliedRead = read
        isHandling = true
        let result = handle(type: .keyDown, event: event)
        isHandling = false
        suppliedRead = nil
        if result != nil { replay(event) }
    }

    private func drain() {
        while !waiting, !isHandling, !pending.isEmpty {
            guard isPowerPointFrontmost() else {
                discardPending("app changed")
                return
            }
            let event = pending.removeFirst()
            if replayVerbatim || isSynthetic(event) {
                replay(event)
                continue
            }
            if needsSelectionRead(event) {
                pending.insert(event, at: 0)
                startRead()
                return
            }
            process(event, read: nil)
        }
        if pending.isEmpty { replayVerbatim = false }
    }

    private func scheduleDrainIfIdle() {
        guard !waiting, !isHandling, !pending.isEmpty else { return }
        DispatchQueue.main.async { [weak self] in self?.drain() }
    }

    /// T runs on `queue` like every other macro. Keys typed meanwhile are held and
    /// replayed as typed once the text box exists, so none reach the slide early.
    private func dispatchTextbox() {
        recordReplayOrigin()
        waiting = true
        replayVerbatimPending = true
        generation += 1
        let token = generation
        armWaitTimer(token, after: textboxLimit)
        fputs("[PPTKKeys] dispatch T, holding keys\n", stderr)
        queue.async { [weak self] in
            self?.dispatch("T")
            DispatchQueue.main.async { self?.textboxFinished(token) }
        }
    }

    private func textboxFinished(_ token: Int) {
        guard token == generation, waiting else { return }
        replayVerbatimPending = false
        endWait()
        guard replayTargetValid() else {
            discardPending("replay target changed")
            return
        }
        if !pending.isEmpty {
            fputs("[PPTKKeys] replaying \(pending.count) held keys\n", stderr)
        }
        replayVerbatim = true
        drain()
    }

    private func recordReplayOrigin() {
        replayOriginPID = NSWorkspace.shared.frontmostApplication?.processIdentifier ?? 0
        replayOriginWindow = focusedWindow(pid: replayOriginPID)
    }

    /// Same PowerPoint process and window as at T, and no dialog or menu in front (Windows ReplayGuard).
    private func replayTargetValid() -> Bool {
        guard isPowerPointFrontmost() else { return false }
        let pid = NSWorkspace.shared.frontmostApplication?.processIdentifier ?? 0
        guard pid == replayOriginPID else { return false }
        if let origin = replayOriginWindow, let now = focusedWindow(pid: pid), !CFEqual(origin, now) {
            return false
        }
        return !focusInsideDialog() && !focusInsideMenu()
    }

    /// A replayed key that never comes back (another tap dropped it) stops counting after a second.
    private func replayStillInFlight() -> Bool {
        guard replaysInFlight > 0 else { return false }
        if Date().timeIntervalSince(lastReplayAt) > replayInFlightLimit {
            replaysInFlight = 0
            return false
        }
        return true
    }

    private func replay(_ event: CGEvent) {
        guard let down = event.copy() else { return }
        down.setIntegerValueField(.eventSourceUserData, value: replayMarker)
        replaysInFlight += 1
        lastReplayAt = Date()
        down.post(tap: .cgSessionEventTap)
        guard let up = event.copy() else { return }
        up.type = .keyUp
        up.setIntegerValueField(.eventSourceUserData, value: replayMarker)
        up.post(tap: .cgSessionEventTap)
    }

    private func discardPending(_ reason: String) {
        replayVerbatim = false
        guard !pending.isEmpty else { return }
        fputs("[PPTKKeys] discarded \(pending.count) held keys (\(reason))\n", stderr)
        pending.removeAll()
    }

    func handle(type: CGEventType, event: CGEvent) -> Unmanaged<CGEvent>? {
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
                if chord == pasteContentToken {
                    return handleBracket(event: event, keycode: keycode, isRepeat: isRepeat, token: chord)
                }
                return handleCanvasToken(event: event, keycode: keycode, isRepeat: isRepeat, token: chord)
            } else if let token = bracketChord(event) {
                dropPrefix(timeout: false)
                return handleBracket(event: event, keycode: keycode, isRepeat: isRepeat, token: token)
            } else if let token = digitChord(event) {
                dropPrefix(timeout: false)
                return handleCanvasToken(event: event, keycode: keycode, isRepeat: isRepeat, token: token)
            } else if optionPeriodEvent(event) {
                dropPrefix(timeout: false)
                return handleBracket(event: event, keycode: keycode, isRepeat: isRepeat, token: "OPT .")
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
            } else if keycode == escapeKeycode, !modified {
                // Windows: Esc cancels the sequence and is swallowed while the canvas is legal.
                let context = readContext(isRepeat: false, letter: "ESC")
                dropPrefix(timeout: false)
                if sequenceAllowed(context) {
                    fputs("[PPTKKeys] sequence cancelled, Esc swallowed\n", stderr)
                    swallowedKeycode = keycode
                    return nil
                }
                swallowedKeycode = nil
                return Unmanaged.passUnretained(event)
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
            if chord == pasteContentToken {
                return handleBracket(event: event, keycode: keycode, isRepeat: isRepeat, token: chord)
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

        if let token = digitChord(event) {
            guard front else {
                swallowedKeycode = nil
                return Unmanaged.passUnretained(event)
            }
            return handleCanvasToken(event: event, keycode: keycode, isRepeat: isRepeat, token: token)
        }

        if optionPeriodEvent(event) {
            guard front else {
                swallowedKeycode = nil
                return Unmanaged.passUnretained(event)
            }
            return handleBracket(event: event, keycode: keycode, isRepeat: isRepeat, token: "OPT .")
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
        // Windows DecideSaveSlides: thumbnail pane with slides selected only (not the sorter, not the slide pane).
        guard context.show == .editing, context.gate == .slides, context.pane == .thumbnails, context.surface != .blocked else { return false }
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
                dispatchTextbox()
            } else {
                queue.async { [weak self] in self?.dispatch(token) }
            }
        }
        swallowedKeycode = keycode
        return nil
    }

    private func readContext(isRepeat: Bool, letter: String) -> KeyContext {
        if isRepeat, repeatGate != .unknown {
            return KeyContext(gate: repeatGate, show: .editing, surface: surfaceFor(repeatGate), detail: "repeat", pane: repeatPane)
        }
        // Never ask PowerPoint on the tap thread. The read was made on `queue` before this key was handled.
        let read = suppliedRead ?? (.unknown, .unknown, "selection=not read", .unknown)
        repeatGate = read.gate
        repeatPane = read.pane
        let context = KeyContext(gate: read.gate, show: read.show, surface: surfaceFor(read.gate), detail: read.detail, pane: read.pane)
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
    private func modeRead() -> ModeRead {
        switch appleScript(modeScript) {
        case .failed(let number, let message):
            if number == -1743, !loggedAutomation {
                loggedAutomation = true
                fputs("Automation: not granted. The key is not intercepted. Allow this program to control Microsoft PowerPoint.\n", stderr)
            } else if !loggedReadFailure {
                loggedReadFailure = true
                fputs("Selection read failed \(number) \(message). The key is not intercepted.\n", stderr)
            }
            return (.unknown, .unknown, "selection=unknown", .unknown)
        case .value(let text) where text == "show":
            return (.notCanvas, .slideshow, "slideshow", .blocked)
        case .value(let text):
            return classify(text)
        }
    }

    /// Windows ShortcutGate.Classify: text wins; masters, notes page, outline and print preview block;
    /// thumbnails and the sorter are the slides surface; the slide pane is the canvas unless slides are selected.
    private func classify(_ text: String) -> ModeRead {
        let parts = text.components(separatedBy: "|")
        let view = parts.count > 1 ? parts[0] : "unknown"
        let selection = parts.count > 1 ? parts[1] : text
        let detail = "view=\(view) selection=\(selection)"
        if selection.contains("text") { return (.editing, .editing, detail, paneKind(view)) }
        let pane = paneKind(view)
        switch pane {
        case .blocked:
            return (.notCanvas, .editing, detail, pane)
        case .thumbnails, .sorter:
            return (.slides, .editing, detail, pane)
        case .slide, .unknown:
            if selection.contains("slide") { return (.slides, .editing, detail, pane) }
            if selection.contains("shape") || selection.contains("none") { return (.canvas, .editing, detail, pane) }
            return (.unknown, .unknown, detail, pane)
        }
    }

    private func paneKind(_ view: String) -> PaneKind {
        switch view {
        case "slide view", "normal view": return .slide
        case "thumbnail view": return .thumbnails
        case "slide sorter view": return .sorter
        case "unknown": return .unknown
        default: return .blocked
        }
    }

    private func dispatch(_ token: String) {
        let parameter = appleScriptParameter(token)
        let source = """
        tell application "Microsoft PowerPoint"
            run VB macro macro name "\(macroName)" list of parameters {"\(parameter)"}
        end tell
        """
        defer {
            // Windows DispatchPacing: grid commands crashed Office when run back to back.
            if gridTokens.contains(token) { usleep(gridPacingMicroseconds) }
        }
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
    // Windows KeyLayout: letter keys are never punctuation.
    if keyLetters[Int64(keycode)] != nil { return nil }
    let shift = event.flags.contains(.maskShift)
    guard let text = layoutCharacter(keycode: keycode, shift: shift) else { return nil }
    if text == "/" || text == "." { return text }
    if !shift, let kept = keptUsPunctuation(keycode: keycode, text: text) { return kept }
    return nil
}

/// Windows KeyLayout.KeepUsKey: if the US "." or "/" key types a letter on this layout, it keeps its US meaning.
private func keptUsPunctuation(keycode: UInt16, text: String) -> String? {
    guard let scalar = text.unicodeScalars.first, CharacterSet.letters.contains(scalar) else { return nil }
    if keycode == usPeriodKeycode { return "." }
    if keycode == usSlashKeycode { return "/" }
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

/// Command+Shift+C, Command+Shift+V and Command+Shift+T (Windows Ctrl+Shift+C/V/T). Option and Control stay with macOS.
/// The letter is the same physical key as C, V and T, so Shift does not turn it into another character.
private func formatChord(_ event: CGEvent) -> String? {
    let flags = event.flags.intersection([.maskCommand, .maskShift, .maskAlternate, .maskControl])
    guard flags == [.maskCommand, .maskShift] else { return nil }
    let keycode = event.getIntegerValueField(.keyboardEventKeycode)
    guard let letter = keyLetters[keycode], letter == "C" || letter == "V" || letter == "T" else { return nil }
    return "CMD SHIFT \(letter)"
}

/// Physical K key (ANSI 0x28): not in keyLetters, so a plain K never costs a context read.
private let smartBarKeycode: Int64 = 0x28

/// Paste Content Only also works while editing text (Windows: canvas or text pane), like Option+.
private let pasteContentToken = "CMD SHIFT T"

/// Grid and TK chords on the physical 2-7 keys, so every layout gives the same key.
/// Control alone: CTRL 2...7 (Windows Ctrl+2...7). Option alone: OPT 2 and OPT 3 (Windows Alt+2/3).
/// Command and Shift stay with PowerPoint (Command+2...5 switch views on Mac).
private func digitChord(_ event: CGEvent) -> String? {
    let flags = event.flags.intersection([.maskCommand, .maskShift, .maskAlternate, .maskControl])
    let keycode = event.getIntegerValueField(.keyboardEventKeycode)
    // Control+K opens the Smart Bar (Windows Ctrl+K / Ctrl+Space; Command+K is Hyperlink and
    // Control+Space switches the input source on Mac). Canvas only, so Control+K keeps its text meaning.
    if flags == [.maskControl], keycode == smartBarKeycode { return "CTRL K" }
    guard let digit = digitKeycodes[keycode] else { return nil }
    if flags == [.maskControl] { return "CTRL \(digit)" }
    if flags == [.maskAlternate], digit == 2 || digit == 3 { return "OPT \(digit)" }
    return nil
}

/// Option+. only. The key is the one that types "." without Option.
private func optionPeriodEvent(_ event: CGEvent) -> Bool {
    guard event.flags.contains(.maskAlternate) else { return false }
    guard event.flags.intersection([.maskCommand, .maskControl, .maskShift]).isEmpty else { return false }
    let keycode = UInt16(event.getIntegerValueField(.keyboardEventKeycode))
    guard let text = layoutCharacter(keycode: keycode, shift: false) else { return false }
    return text == "." || keptUsPunctuation(keycode: keycode, text: text) == "."
}

/// Command+/ only. Option and Control stay with macOS. Shift counts only when it is how "/" is typed.
private func commandSlashEvent(_ event: CGEvent) -> Bool {
    guard event.flags.contains(.maskCommand) else { return false }
    guard event.flags.intersection([.maskAlternate, .maskControl]).isEmpty else { return false }
    let keycode = UInt16(event.getIntegerValueField(.keyboardEventKeycode))
    let shift = event.flags.contains(.maskShift)
    guard let text = layoutCharacter(keycode: keycode, shift: shift) else { return false }
    return text == "/" || (!shift && keptUsPunctuation(keycode: keycode, text: text) == "/")
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

/// The focused window of a process, with a short messaging timeout.
private func focusedWindow(pid: pid_t) -> AXUIElement? {
    guard AXIsProcessTrusted(), pid > 0 else { return nil }
    let application = AXUIElementCreateApplication(pid)
    AXUIElementSetMessagingTimeout(application, 0.08)
    guard let raw = copyAttribute(application, kAXFocusedWindowAttribute as String) else { return nil }
    return axElement(raw)
}

/// An open menu (menu bar or context menu) owns the keyboard (Windows FocusProbe treats menu mode as pass).
private func focusInsideMenu() -> Bool {
    guard AXIsProcessTrusted(), let app = NSWorkspace.shared.frontmostApplication else { return false }
    let application = AXUIElementCreateApplication(app.processIdentifier)
    AXUIElementSetMessagingTimeout(application, 0.05)
    guard let raw = copyAttribute(application, kAXFocusedUIElementAttribute as String),
          let element = axElement(raw) else { return false }
    AXUIElementSetMessagingTimeout(element, 0.05)
    let role = attribute(element, kAXRoleAttribute) ?? ""
    return role == "AXMenu" || role == "AXMenuItem" || role == "AXMenuBarItem"
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
    let result = helper.intercept(type: type, event: event)
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
    set paneView to "unknown"
    try
        set paneView to (pane view type of active pane of active window) as text
    end try
    try
        return paneView & "|" & ((selection type of selection of active window) as text)
    on error
        return paneView & "|unknown"
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

// MARK: - Packaged install (PPTK.pkg)

/// The package installs the add-in into Office's system-wide Startup folder
/// (/Library/Application Support/Microsoft/Office365/User Content.localized/Startup/PowerPoint), where PowerPoint
/// loads it at start for every user, with no registration, no file-access dialog and no per-user copy (tested
/// 2026-11-22). The package also leaves the Email Slides AppleScript plugin in its own folder; Office runs it only
/// from each user's PowerPoint scripts folder, so the helper copies it there. The PPTK_* variables are for tests.
private let deploymentSystemRoot = ProcessInfo.processInfo.environment["PPTK_SYSTEM_ROOT"]
    ?? "/Library/Application Support/PPTK"
private let powerPointScriptsFolder = ProcessInfo.processInfo.environment["PPTK_APPLESCRIPTS_DIR"]
    ?? (NSHomeDirectory() + "/Library/Application Scripts/com.microsoft.Powerpoint")

private func syncPowerPointPlugin() -> String {
    let fm = FileManager.default
    let plugin = deploymentSystemRoot + "/InstrumentaAppleScriptPlugin.applescript"
    guard let data = fm.contents(atPath: plugin) else { return "not deployed" }
    let target = powerPointScriptsFolder + "/InstrumentaAppleScriptPlugin.applescript"
    if fm.contents(atPath: target) == data { return "current" }
    do {
        try fm.createDirectory(atPath: powerPointScriptsFolder, withIntermediateDirectories: true)
        try data.write(to: URL(fileURLWithPath: target), options: .atomic)
        return "installed"
    } catch {
        return "error: \(error.localizedDescription)"
    }
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

// The package's LaunchAgent is system-wide, so it cannot name a per-user log path: the helper opens its own log
// (~/Library/Logs/PPTK/PPTKKeys.log, renamed to .1 above 5 MB) when PPTK_LOG_TO_USER_LIBRARY=1.
if ProcessInfo.processInfo.environment["PPTK_LOG_TO_USER_LIBRARY"] == "1" {
    let folder = NSHomeDirectory() + "/Library/Logs/PPTK"
    let log = folder + "/PPTKKeys.log"
    try? FileManager.default.createDirectory(atPath: folder, withIntermediateDirectories: true)
    if let size = (try? FileManager.default.attributesOfItem(atPath: log))?[.size] as? Int, size > 5_000_000 {
        try? FileManager.default.removeItem(atPath: log + ".1")
        try? FileManager.default.moveItem(atPath: log, toPath: log + ".1")
    }
    freopen(log, "a", stdout)
    freopen(log, "a", stderr)
}
// Line-buffered stdout: under launchd the output goes to a log file, which would otherwise be block-buffered.
setvbuf(stdout, nil, _IOLBF, 0)

private let helper = Helper()

switch CommandLine.arguments.dropFirst().first {
case "--describe-focus":
    exit(describeFocus())
case "--version":
    let version = Bundle.main.object(forInfoDictionaryKey: "CFBundleShortVersionString") as? String ?? "development"
    fputs("PPTKKeys \(version)\n", stdout)
    exit(0)
case "--sync-plugin":
    fputs("\(syncPowerPointPlugin())\n", stdout)
    exit(0)
case "--help":
    fputs("Usage: PPTKKeys [--describe-focus | --version | --sync-plugin]\nCtrl+C stops the listener.\n", stdout)
    exit(0)
default:
    exit(helper.start())
}
