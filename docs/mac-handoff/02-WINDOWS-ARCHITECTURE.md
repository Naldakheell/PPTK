# 02 — Windows architecture and the reasons behind it

## 1. The big picture

```
┌──────────────────────── PowerPoint (POWERPNT.EXE) ────────────────────────┐
│  PPTK.ppam  (VBA add-in, AutoLoad via HKCU AddIns key)                     │
│   ├─ Ribbon: src/PPTK/Ribbon/CustomUI.xml → PPTK_RibbonShortcut(tag)       │
│   │                                       or direct public Subs            │
│   ├─ PPTK_DispatchShortcut(token)  (src/Modules/ModulePPTKShortcuts.bas)   │
│   └─ feature modules: ModulePPTK*.bas, forms, classes                      │
└───────────────▲────────────────────────────────────────────────────────────┘
                │ COM Application.Run("PPTK_DispatchShortcut", token)
┌───────────────┴──────────── PPTKKeys helper (dotnet.exe exec PPTKKeys.dll) ┐
│  WH_KEYBOARD_LL hook → FocusProbe (Win32) + COM snapshot → ShortcutGate     │
│  → token queue → message loop → PowerPointSession.Run                       │
│  + pending update watcher, update notice, handoff (production config only) │
└──────────────────────────────────────────────────────────────────────────────┘
 started at sign-in by Task Scheduler "\PPTKKeys Startup" (per user, no admin)
```

**Why two parts:**

- VBA cannot capture global keystrokes or bare letter keys reliably.
- VBA also cannot call Win32: company ASR rule `92E97FA1` blocks `Declare` in Office macros.

So a separate helper captures keys and decides context, and VBA does the PowerPoint work. The **token contract** (`PPTK_DispatchShortcut(token)`) is the only interface between the two. The macOS project already follows the same split (Swift helper → AppleScript `run VB macro`), which validates the design.

## 2. The PowerPoint add-in (VBA)

| Responsibility | Module(s) | Notes |
|---|---|---|
| Dispatcher and small commands | `ModulePPTKShortcuts.bas` (~2100 lines) | `PPTK_DispatchShortcut`, `PPTK_NormalizeToken`; inserts, text tools, enclose icons, save/email slides, `PPTK_ExecuteMso` |
| Ribbon callbacks | `ModulePPTKRibbon.bas` | `PPTK_RibbonShortcut(control)` passes `control.Tag` to the dispatcher (with a `Nothing` guard), plus a few direct callbacks |
| Version and form styling | `src/PPTK/Brand/ModulePPTKBrand.bas` | `PPTK_VERSION`, `PPTK_GetVersion()` (checked by the build), `PPTK_StyleForm` and button helpers |
| Feature modules | `ModulePPTKFixColors`, `ModulePPTKReplaceColors`, `ModulePPTKFlip`, `ModulePPTKGrid`, `ModulePPTKSmartAlign`, `ModulePPTKSmartBar`, `ModulePPTKSlideLibrary`, `ModulePPTKTakeaway` | one feature each (see 01) |
| Inherited geometry | `ModuleObjects`, `ModuleObjectsAlignAndDistribute`, `ModuleObjectsSizeAndPosition` | Instrumenta code, modified for Windows: `SetRealSize`, child-range guards, `PPTK_SwapZOrder` |

### Selection and shape operations (common patterns)

- **Selection:**
  - Read `ActiveWindow.Selection`.
  - Prefer `ChildShapeRange` when the user selected inside a group (fixed in the distribute and align code).
  - Many commands require a shape selection and exit silently otherwise. User errors get `MsgBox` titled "PPTK".
- **Rotated shapes:** `SetRealSize` (`ModuleObjects.bas`) sets the *visible* width and height of a rotated shape while keeping its center. It returns False when the target is geometrically unreachable (about 45° with unequal sides).
  - The Mac branch has an **independently written** `SetRealSize` with a different signature. Compare them before merging.
- **Undo:** `Application.StartNewUndoEntry` groups one command into one undo step.
  - Used in: Grid, Flip, Replace Colors, Match Corner, Unify Type, Remove Margins.
  - Not used in: Fix Colors, Smart Align, TK, Enclose, Paste Content Only, Merge, inserts.
- **State on shapes:** tags (`PPTK_SHORTCUT`, `PPTK_ENCLOSURE_*`) rather than names, so state survives renames and copy/paste.
- **Data safety:** destructive bulk commands take a file backup first (Flip, Replace Colors), or keep hidden recovery copies (Grid delete).
- **Error handling:**
  - `On Error GoTo Clean` with a silent exit.
  - `On Error Resume Next` only around per-object writes.
  - Save and Email restore `DisplayAlerts` and the `Saved` flag, and treat cancel errors (1004, 4198, 18) as normal.

### Linked Groups and shape replacement

**Neither exists** in the repository (see 01 §D). Do not plan Mac work around them as existing Windows features.

## 3. The PPTKKeys helper (C#, .NET 10, `src/PPTKKeys/windows/`)

### Threads and loop

- One STA thread runs a message-only window (`PPTKKeys.Message.<pid>`).
- **Hooks:**
  - `WH_KEYBOARD_LL`;
  - `SetWinEventHook` for foreground changes;
  - a per-PowerPoint-process `EVENT_OBJECT_FOCUS` hook.
- **Timers:**
  - A 1 s refresh timer re-attaches COM and refreshes the layout tables. It also drives the update watcher and the notice.
  - A 1 s sequence timer.
- **Private messages:** `WmDispatch`, `WmRefresh`, `WmRunOnHelperThread`.

### Inside the hook (`HelperApp.OnHook`): fast, no COM

1. **Keys always passed straight through:**
   - injected keys (`LLKHF_INJECTED`), including the helper's own replay;
   - modifier keys;
   - keys that are not relevant.
2. **Repeats and key-ups:** swallow the auto-repeats and the key-up of the key that ran a command.
3. **Context:**
   - `FocusProbe.Inspect()` uses Win32 only: the foreground process image, the frame class `PPTFrameClass`, and the focus chain must contain `paneClassDC`/`mdiClass`. Edit, RichEdit, NetUI, dialog and menu classes block.
   - The cached COM `Snapshot` gives the surface.
4. **Decision:** sequence, then `\`, then chord, then icon, then bare letter (`ShortcutGate`, `ShortcutSequence`, `ShortcutChords`).
5. **Dispatch:** post the token to the queue and swallow the key. If anything fails, the key passes (**fail open**).

**Why:**

- A low-level hook must return fast, or Windows removes it (`LowLevelHooksTimeout`). So the hook never calls COM.
- COM runs later, on the same thread's message loop.
- The context comes from two sources:
  - Win32 window classes, which are immediate but coarse;
  - the COM selection type, which is precise but asynchronous.

  Neither is enough alone: focus can sit in the ribbon while the selection still says "shapes".

### PowerPoint session (`PowerPointSession.cs`)

- **Attach:** `GetActiveObject` + interop events: `WindowSelectionChange`, `SlideShowBegin/End`, `WindowActivate`.
- **Read:** `ShortcutGate.Classify(slideshow, pane, selection)` gives Canvas / Slides / Text / Blocked / Unknown.
- **Invalidate:** every focus or foreground change sets Unknown (pass) until a refresh lands. This avoids acting on a stale "Canvas" after the user clicked into a text box.
- **Disconnect:** RPC HRESULTs `0x800706BA`, `0x80010108`, `0x800706BE` cause a `Drop()`, and the timer re-attaches.
- **T text hold (`ShortcutGate.Reduce`):** after T, bare letters stay off until PowerPoint reports the text selection, so a stale refresh cannot re-enable shortcuts while the user is typing into the new box.

### Layouts

- Letters use virtual keys (positional).
- Punctuation comes from the active layout (`KeyLayout.Build`, `ToUnicodeEx`). The tables are rebuilt when the foreground thread's layout changes.
- **Arabic rule (`LayoutRule.KeepsUsPunctuation`):** if the US punctuation key types a letter in this layout (period → ز), it keeps its US meaning.
- **AltGr:** Windows sends a synthetic Left Ctrl with AltGr. `AltGrRule` treats Right Alt as Alt only when the layout types nothing on that key.

### Text replay after T

- Characters typed while the textbox is being created are queued (max 128) and replayed with `SendInput` Unicode.
- **Why `ReplayGuard`:** replay goes only to the same PowerPoint process and frame window, and only if no dialog or menu opened. Otherwise the text is discarded. It must never be typed into another app.

### Single instance

- **Mutex:** `Local\PPTKKeys.Windows`.
- **Console:** the scheduled task starts `dotnet.exe`, a console program, so Windows creates a console window. The helper calls `FreeConsole` when it is the only process attached (`ReleaseOwnConsole`), so the window disappears. It is kept when started from a terminal.

## 4. State persistence

| What | Where | Owner |
|---|---|---|
| Feature settings | `HKCU\Software\VB and VBA Program Settings\PPTK\...` (`SaveSetting`) | VBA |
| Shape state | shape tags | VBA |
| Installation state | `%LOCALAPPDATA%\PPTK\state.json` (+ `state.lock`, `transaction.json`) | installer + helper engine |
| Installed version | `HKCU\Software\PPTK\InstalledVersion` | installer + engine |
| Add-in registration | `HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK` | installer + engine |
| Logs | `%LOCALAPPDATA%\PPTK\logs\` | installer, helper (production config only) |

## 5. Notification and update mechanisms

**Pending update:**

- If PowerPoint has the add-in open, the installer stages the new version (`versions\<v>` + `pending\`) and exits 1618.
- `PendingUpdateWatcher` (every 5 s, on a background thread) runs `PendingAddinUpdate` once no PowerPoint process is left in the session.
- It applies the update in this order: journal, task, add-in, state, registry, then commits by deleting the journal.
- On failure it undoes everything. Hidden PowerPoint processes are waited for (30 s threshold), never killed.

**Notice:**

- `UpdateNoticeController.Tick` runs on the message thread every second and reads only cached state.
- The window is `NotificationWindow`: `WS_EX_NOACTIVATE | WS_EX_TOPMOST | WS_EX_TOOLWINDOW`, `SW_SHOWNOACTIVATE`, `MA_NOACTIVATE`, closes after 10 s or on click.
- The "shown" flag is written off the message thread, under `state.lock`.
- **Why live window checks:** the cached COM snapshot said "slideshow" for 30 s after the show ended, because it refreshes only on PowerPoint events. A live check for visible `screenClass`/`PodiumParent` windows fixed it (verified on a real device).

**Handoff after an update that changed the helper:**

- The old helper starts the new one with `--handoff <pid> <nonce>`.
- Named events: Prepared → Released → Active (Abort on failure).
- **Invariant:** hook ownership = mutex ownership. A process installs the hook only while it owns the mutex, and removes it before releasing.
- If the new helper fails, the old one takes the hook back.
- Forced kill of a stuck new helper is disabled.

## 6. Error handling and recovery principles

1. **Fail open for keys:** when unsure, the key goes to PowerPoint. A missed shortcut is cheap; a swallowed keystroke while typing is not.
2. **Fail safe for files:** every install, update, uninstall and rollback is journaled and undone on failure. The add-in is never replaced while PowerPoint can hold it.
3. **Never kill user processes:** PowerPoint is never closed. The helper is only asked to close.
4. **Test names in tests:** live tests use `PPTK-Test-<id>` tasks, keys and mutexes. Production names are refused in test mode.

## 7. Why the final approaches won over alternatives

| Decision | Alternatives rejected | Reason |
|---|---|---|
| `dotnet.exe exec PPTKKeys.dll` | Unsigned apphost `PPTKKeys.exe`; self-contained single-file exe | ASR rule `01443614` blocked `PPTKKeys.exe` (event log evidence). Microsoft-signed `dotnet.exe` was never blocked. |
| Private official .NET runtime (1.1.0) | Machine-wide .NET dependency in Intune | One package for x64 + ARM64; no admin; no separate dependency app. |
| VBA without `Declare` | Win32 calls from VBA (e.g. focus, timers) | ASR rule `92E97FA1` blocks them; the Smart Bar uses Hide/Show and a custom scrollbar instead of the mouse wheel. |
| COM `Application.Run` with a token | SendKeys/KeyTips; UI Automation clicks | Deterministic; one macro entry point; testable token list. |
| Context = Win32 focus + COM selection | Either source alone | See §3. |
| Live slideshow check for the notice | Cached COM snapshot | The snapshot was stale for 30 s after a show ended. |
| Stage + apply-on-close updates | Replace the add-in while PowerPoint runs; close PowerPoint | The file is locked or unsafe; closing PowerPoint is unacceptable to users. |
