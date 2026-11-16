# 03 — Reusable components

There are three categories:

- **A:** reusable on macOS as-is (or nearly).
- **B:** reusable concepts or algorithms that need platform adaptation.
- **C:** Windows-specific code; do not port it.

Ground rule: **the Mac project already works for its feature set** (`origin/mac-development`, `4503753`). Reuse means adding Windows features into the existing Mac structure, not replacing Mac code.

## A. Directly reusable (VBA that runs in PowerPoint for Mac)

VBA is shared between the platforms, but only code that avoids Windows-only objects. Check each candidate for these **Windows-only dependencies**:

- `Dir$` with Unicode paths;
- `Environ("USERPROFILE")`, `%TEMP%`;
- `GetObject("New:{…}")` (MSForms DataObject);
- `Outlook.Application` COM;
- `FileSystemObject` (not available on Mac);
- `ExecuteMso` IDs that differ on Mac (e.g. `IconInsertFromFile` vs `IconsPane`);
- MSForms features that behave differently on Mac (UserForms are limited on Mac Office; **verify each form**).

| Component | File / functions | Mac reuse notes |
|---|---|---|
| Shape inserts with nudge | `ModulePPTKShortcuts.bas`: `PPTK_InsertShape`, `PPTK_PlaceOrigin`, `PPTK_TokenCovers` | **Already on Mac** in the same form. Keep the Mac copy. |
| Match corner roundness | `PPTK_MatchCornerRoundness` | Pure object model. Likely portable as-is. |
| Unify shape type | `PPTK_UnifyShapeType`, `PPTK_AssignAutoShapeType` | Pure object model. Likely portable. |
| Merge text, swap text | `PPTK_MergeText`, `PPTK_SwapText` | Mac uses the Instrumenta `ObjectsTextMerge` / `ObjectsSwapTextNoFormatting`. Compare behavior before switching; do not switch only for uniformity. |
| Remove margins + fit | `PPTK_ZeroMargins` | Object model; check `TextFrame2` autosize behavior on Mac. |
| Enclose icons | `PPTK_EncloseIcons` and helpers | Object model + tags. Check `msoGraphic` detection on Mac (icons differ there per the Instrumenta README). |
| Smart Align | `ModulePPTKSmartAlign.bas` (+ `PPTKMatrixForm`, `PPTKCentralForm`) | The algorithms are pure geometry. The forms need Mac UserForm testing. |
| Grid rows/columns | `ModulePPTKGrid.bas` (`PPTK_Recognize` and others) | Pure geometry + object model. The largest win; test hidden recovery copies on Mac. |
| TK placeholders | `ModulePPTKTakeaway.bas` | Object model only. |
| Fix Slide Colors | `ModulePPTKFixColors.bas` | Object model; the spec is `ThemeHarmonize.cs` (see B). |
| Flip slide/deck | `ModulePPTKFlip.bas` | Object model. The backup path uses Windows `Documents` folder logic; adapt it. |
| Find & Replace Colors | `ModulePPTKReplaceColors.bas` | Walker and matching are portable. `PPTKColorsForm` needs Mac form testing. The backup reuses the Flip helper. |
| `SetRealSize` | `ModuleObjects.bas` | Both branches have one, written independently. **Pick one after comparing tests**; do not keep two. |
| Version constant | `ModulePPTKBrand.bas` `PPTK_VERSION`, `PPTK_GetVersion` | Reusable concept. The file itself says "do not import into the Mac workbook" (form styling); split out the version function if needed. |

## B. Reusable concepts and algorithms (adapt)

| Concept | Windows source | How Mac can reuse it |
|---|---|---|
| **Token contract** | `PPTK_DispatchShortcut(token)`, `PPTK_NormalizeToken` | Already shared. Unify the token names (see 05 §1) by accepting both spellings in one dispatcher. |
| **Context gate matrix** | `ShortcutGate.Decide`, `DecideIcon`, `DecideSaveSlides`, `Classify`, `EffectiveSurface` | Port the *rules* into the Swift gate (`directAllowed`, `sequenceAllowed`, …). Use `ShortcutGateTests.cs` as the case list for XCTest. |
| **T text-hold state machine** | `ShortcutGate.Reduce`, `TryQueueText`; `ReplayGuard` | Mac solved "T then type" by synchronous dispatch. Adopt the hold/replay design only if fast typing loses characters on Mac. |
| **Sequence semantics** | `ShortcutSequence` (`Advance`, `Next`, `Match`, `Expired`) | Mac has prefixes with the same 1 s window. Align the details: invalid second key re-processed, Esc swallowed, prefix never typed, `X` and `.` prefixes. |
| **Spacing repeat throttle** | `ShortcutRepeat` (50 ms, stop on release) | Mac has no H±/V± repeat. Port the rules if wanted. |
| **Grid dispatch pacing** | `DispatchPacing` (100 ms after grid tokens) | Keep it if grid commands are added on Mac (the crash was Office-internal). |
| **Layout rules** | `LayoutRule.KeepsUsPunctuation`, `IsLetterKey` | Mac already resolves `/ . + -` by layout and `\ ; [` physically; compare with the Arabic rule. AltGr does not apply. |
| **Theme color mapping** | `ThemeHarmonize.cs` + `ThemeHarmonizeTests.cs` | Specification and thresholds of Fix Slide Colors (`NeutralChroma 0.12`, `MaxBrightness 0.55`, `MinimumContrast 3.0`). |
| **Update notice policy** | `NotificationPolicy`, `NoticeScheduler`, `UpdateNoticeController` | Same rules for a future Mac notice: PowerPoint in front, no slideshow/dialog/menu, 5 s idle, once per version, never take focus (`NSPanel` `.nonactivatingPanel`). |
| **Pending update on close** | `PendingAddinUpdate.cs`, `JOURNAL.md` | Same flow for a future Mac installer: journal → swap → commit; undo on failure; never close PowerPoint. |
| **Handoff invariant** | `HelperHandoff.cs` | "Only the lock owner may hold the event tap" — applicable if a Mac helper is ever replaced while running. |
| **Package manifest** | `version.json` + hashes, `detect.ps1` checks | Same manifest idea for a Mac `.pkg`. |

## C. Windows-specific: do NOT port

| Code | Why |
|---|---|
| `Program.cs` hook, message loop, timers, `NativeMethods.cs` | Win32 (`WH_KEYBOARD_LL`, `SetWinEventHook`, `SetTimer`). Mac already has CGEventTap + run loop. |
| `FocusProbe.cs` | Win32 window classes. Mac uses AX roles and subroles (`focusInsideDialog`, `focusAcceptsTyping`, `surfaceVeto`); keep them. |
| `KeyLayout.cs`, `AltGrRule` | `ToUnicodeEx`, AltGr. Mac uses `UCKeyTranslate`. |
| `PowerPointSession.cs` | COM, interop events, RPC HRESULTs. Mac uses AppleScript. |
| `NotificationWindow.cs`, `PowerPointWindows.cs` | Win32 windows and classes. |
| `LiveDeployment.cs`, `HelperWiring.cs` production config, `src/Deploy/*` | Task Scheduler, HKCU, PowerShell, Authenticode, Intune. |
| VBA: `PPTK_EmailSelectedSlides` (Outlook COM), `PPTK_ClipboardText` (MSForms DataObject), `PPTK_InsertIcon` (`IconInsertFromFile`), FSO usage in Save Slides | Windows Office/COM specifics. The Mac branch already has its own Email (AppleScriptTask plugin), Icon (`IconsPane`) and Save Slides (`MacSaveAsDialog`) paths. |
| `ModulePPTKBrand.bas` form styling | Explicitly Windows-only per its header. |
