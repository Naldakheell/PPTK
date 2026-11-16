# 05 — Keyboard shortcuts: Windows behavior and Mac parity

Sources (verified against code on `windows-development` at `49fe2f2`):

- Windows helper: `src/PPTKKeys/windows/ShortcutGate.cs`, `Program.cs` (`HelperApp.OnHook`), `KeyLayout.cs`, `FocusProbe.cs`, `PowerPointSession.cs`.
- Windows VBA dispatcher: `src/Modules/ModulePPTKShortcuts.bas` → `PPTK_DispatchShortcut(token)`.
- Mac helper: `origin/mac-development:src/PPTKKeys/mac/Sources/PPTKKeys/main.swift` (commit `4503753 mac1`).
- Mac dispatcher: `origin/mac-development:src/Modules/ModulePPTKShortcuts.bas`.
- Cross-check: the employee cheat sheet `docs/cheat-sheet/` agrees with the Windows code; no extra or missing keys.

## 1. Shared contract: helper → token → VBA

Both platforms already use the same architecture. Keep it.

```
key event ──► helper decides (context gate) ──► token string ──► PowerPoint VBA
                                                   "A L"          PPTK_DispatchShortcut(token)
```

- **Windows transport:** COM `Application.Run("PPTK_DispatchShortcut", token)` in `PowerPointSession.Run`.
- **Mac transport:** in-process `NSAppleScript`:
  ```
  tell application "Microsoft PowerPoint" to run VB macro macro name "PPTK_DispatchShortcut" list of parameters {"<token>"}
  ```
  It is serialized by `scriptLock`.
- **VBA side:**
  - `PPTK_NormalizeToken` uppercases the token and collapses whitespace.
  - A `Select Case` runs the macro.
  - An unknown token does nothing (`Case Else`).
- **Rule to keep:** the helper decides *whether* a key is a PPTK shortcut. VBA decides *what* to do. No PowerPoint logic lives in the helper.

### Token naming differences (must be reconciled)

The same feature uses different tokens on each platform. A Mac token sent to the Windows `.bas` falls into `Case Else` and does nothing (and the reverse).

| Feature | Mac token | Windows token |
|---|---|---|
| Swap text | `CMD /` | `CTRL /` |
| Copy formatting | `CMD SHIFT C` | `CTRL SHIFT C` |
| Paste formatting | `CMD SHIFT V` | `CTRL SHIFT V` |
| Zero margins | `OPT .` | `ALT .` |
| Font size + / − | `FONT +` / `FONT -` | `CTRL ]` / `CTRL [` |
| Bring forward / send backward | `ORDER +` / `ORDER -` | `CTRL SHIFT ]` / `CTRL SHIFT [` |

**Recommendation (not implemented):**

- Prefer **semantic tokens** (`FONT +`, `ORDER +`, `SWAP TEXT`, …) in one shared dispatcher that accepts both spellings.
- Do not rename the Mac helper's tokens until the shared `.bas` accepts them. Today each platform ships its own `.bas`.

## 2. Windows shortcut inventory (exact)

### Context terms

- **Canvas:**
  - `FocusProbe` reports `FocusKind.Pane`;
  - **and** the COM snapshot surface is `Canvas`: the slide pane (`ViewSlide`) with nothing or with shapes selected;
  - **and** no slideshow is running and no text hold is active.
- **Slides:** thumbnails, the sorter, or the slide pane with slides selected.
- **Text:** a text selection (editing inside a shape).
- **Fails open:** anything Unknown (focus changing, PowerPoint not attached, read failed) passes the key to PowerPoint.

### 2.1 Bare letters

No Ctrl, Shift, Alt or Win. Letters match by virtual key (physical position on any layout).

| Key | Token | VBA | Gate | Status |
|---|---|---|---|---|
| T | `T` | `PPTK_InsertTextbox`, then types the captured text | canvas | tested (unit + real use) |
| R | `R` | `PPTK_InsertRectangle` | canvas | tested |
| U | `U` | `PPTK_InsertRoundedRectangle` | canvas | tested |
| O | `O` | `PPTK_InsertOval` | canvas | tested |
| C | `C` | `PPTK_InsertChevron` | canvas | tested |
| P | `P` | `PPTK_InsertPentagon` | canvas | tested |
| G | `G` | `PPTK_InsertTriangle` | canvas | tested |
| M | `M` | `PPTK_InsertParallelogram` | canvas | tested |
| N | `N` | `PPTK_InsertNewSlide` | canvas **or** slides | tested |
| E | `E` | `PPTK_EncloseIcons` | canvas | Windows-only |
| Q | `Q` | `PPTK_FixSlideColors` | canvas | Windows-only |
| I | `I` | `PPTK_InsertIcon` (Windows Insert Icon dialog, on its own STA thread) | canvas (`DecideIcon`) | tested; Mac uses `ExecuteMso "IconsPane"` |
| `\` | `\` | `PPTK_SaveSelectedSlides` | focus Pane, **thumbnail pane** with slides selected (not the sorter, not the canvas); bare only | tested |

"Tested" means covered by the pure gate tests (`ShortcutGateTests`, `SequenceTests`) and used in manual sessions. There is no automated end-to-end test with real PowerPoint.

### 2.2 Two-key sequences

The prefix, then the second key, within **1000 ms** (`ShortcutSequence.TimeoutMs`).

| Sequence | VBA | Mac |
|---|---|---|
| A L / A R / A T / A B / A M / A C | `ObjectsAlignLefts/Rights/Tops/Bottoms/Middles/Centers` | same |
| A A | `PPTK_MatchCornerRoundness` | missing |
| A P / A X / A Z | `PPTK_AlignProcess` / `PPTK_AlignMatrix` / `PPTK_AlignCentral` (Smart Align) | missing |
| S W / S H / S S | `ObjectsSameWidth` / `SameHeight` / `SameHeightAndWidth` | same (Mac uses its own `SetRealSize`) |
| D H / D V | `ObjectsDistributeHorizontally` / `Vertically` | same |
| H + / H − | `ObjectsIncrease/DecreaseSpacingHorizontal`; **hold to repeat** | same (no repeat on Mac) |
| V + / V − | `ObjectsIncrease/DecreaseSpacingVertical`; **hold to repeat** | same |
| X T | `PPTK_UnifyShapeType` | missing |
| . T | `PPTK_MergeText` | same feature, via Instrumenta `ObjectsTextMerge` |

Sequence rules (`ShortcutSequence`):

- A prefix arms only on the canvas, with no Ctrl, Alt or Win.
- Shift+letter does not arm.
- The period prefix is the key that types `.` on the current layout. Shift is allowed if that layout needs it.
- Second key:
  - Letters must be unshifted.
  - `+` is Shift+=, numpad +, or any key whose layout character is `+`.
  - `−` is unshifted minus or numpad −; Shift+minus (underscore) does not count.
- The second key does **not** require releasing the prefix. While the prefix is held, its auto-repeats are swallowed.
- Esc cancels, and the Esc is swallowed.
- An invalid second key cancels the sequence and is processed as a new key: `A` then `U` inserts a rounded rectangle; `D` then `D` re-arms.
- The prefix key itself is never typed, even when the sequence is cancelled.
- Leaving PowerPoint, or focus moving off the pane, cancels the sequence.

### 2.3 Chords

Not with Win, and not with Ctrl+Alt together (so real AltGr is never a chord).

| Keys | Token | VBA | Gate |
|---|---|---|---|
| `/` (layout slash) | `/` | `ObjectsSwapPosition` | canvas |
| Ctrl+/ | `CTRL /` | `PPTK_SwapText` | canvas |
| Ctrl+Shift+C / V | `CTRL SHIFT C` / `V` | `PPTK_CopyFormatting` / `PPTK_PasteFormatting` | canvas |
| Ctrl+Shift+T | `CTRL SHIFT T` | `PPTK_PasteContentOnly` | canvas **or text editing** |
| Alt+. | `ALT .` | `PPTK_ZeroMargins` | canvas **or text editing** |
| Ctrl+] / Ctrl+[ | `CTRL ]` / `CTRL [` | `ExecuteMso "FontSizeIncrease"` / `"FontSizeDecrease"` | canvas |
| Ctrl+Shift+] / Ctrl+Shift+[ | `CTRL SHIFT ]` / `[` | `ExecuteMso "ObjectBringForward"` / `"ObjectSendBackward"` | canvas |
| Ctrl+2 / Ctrl+3 | `CTRL 2` / `CTRL 3` | `PPTK_AddRightColumn` / `PPTK_AddBottomRow` (grid) | canvas |
| Alt+2 / Alt+3 | `ALT 2` / `ALT 3` | `PPTK_DeleteRightColumn` / `PPTK_DeleteBottomRow` | canvas |
| Ctrl+4 / 5 / 6 / 7 | `CTRL 4`–`CTRL 7` | `PPTK_AddRightTK` / `AddBottomTK` / `AddTopTK` / `AddLeftTK` | canvas |
| Ctrl+K, Ctrl+Space | `CTRL K`, `CTRL SPACE` | `PPTK_ShowSmartBar` | canvas |

Notes:

- The digit chords use the top row or numpad keys 2–7. Shift blocks them.
- Grid tokens are paced 100 ms apart (`DispatchPacing.GapMs`). This came from a crash in `mso20win32client.dll` when Ctrl+2 ran 0 ms after Alt+2.
- The VBA also accepts these tokens, which are sent **only** from the Smart Bar and the Ribbon, never from keys (`IsRunnable` rejects them from the helper): `CMD ADD COLUMN EXPAND`, `CMD ADD ROW EXPAND`, `CMD DELETE COLUMN SHRINK`, `CMD DELETE ROW SHRINK`, `CMD FLIP SLIDE`, `CMD FLIP PRESENTATION`, `CMD REPLACE COLORS`.

### 2.4 Mac-only shortcuts (exist on Mac, deliberately not on Windows)

| Mac key | Token | VBA |
|---|---|---|
| `;` | `;` | `ShowSlideLibrary` (Instrumenta slide library) |
| `[` (bare) | `EMAIL SLIDES` | `EmailSelectedSlides` (via `AppleScriptTask` Outlook plugin) |

The Windows `ModulePPTKShortcuts.bas` header says Slide Library and Email Slides are **not** shortcuts on Windows; they are Ribbon features there. Decide before porting whether Mac keeps them. Do not remove them silently.

## 3. Shared semantics to keep identical on Mac

1. **Same keys mean the same action.** Use the modifier mapping Ctrl→Cmd and Alt→Option already chosen on Mac. Bare letters and sequences stay identical.
2. **Context gate:**
   - Act only on the slide canvas. `N` and `\` also act on slide selection.
   - Never act while editing text, except paste-content-only, zero margins and the font-size chords. On Mac the font-size chords also work in text; on Windows they are canvas-only.
   - Never act in dialogs, menus, the ribbon, the Notes pane, a slideshow or another app.
3. **Fail open.** If the context is uncertain, pass the key to PowerPoint. Both platforms already do this.
4. **Sequence timing:** 1 s window. Esc cancels. The prefix is never typed. An invalid second key restarts as a fresh key.
5. **Auto-repeat:** never repeat a one-shot command. Only H± and V± repeat (Windows: min 50 ms between runs, stops on release).
6. **T then typing:** characters typed right after `T` must land in the new text box, not be lost and not become shortcuts (see §4).

## 4. Edge cases (and how Windows handles them)

| Case | Windows handling | Mac status |
|---|---|---|
| **Text editing** | Text selection gives `Surface.Text`; keys pass. RichEdit and ribbon edit boxes are `RibbonOrEdit` (FocusProbe). | Mac uses AX role walks (`focusAcceptsTyping`) plus the AppleScript `selection type`. |
| **T then fast typing** | After `T`, a text hold (`ShortcutGate.Reduce`) turns shortcuts off until PowerPoint reports a text selection. Letters, digits and space typed meanwhile are queued (max 128) and replayed with `SendInput` Unicode after the macro, only to the same process and frame (`ReplayGuard`). | Mac dispatches `T` **synchronously** on the tap thread instead, so the next keys arrive after the textbox exists. Edge cases with very fast typing are unverified on Mac. |
| **Dialogs and menus** | Foreground class not `PPTFrameClass`, or menu mode (`GetGUIThreadInfo`), or `#32770`/`NUIDialog`/`bosa_sdm*`: pass. | AX: AXSheet, AXDialog, AXSystemDialog: pass. |
| **PowerPoint focus** | Foreground pid image `POWERPNT.EXE`; focus hwnd chain must contain `paneClassDC` or `mdiClass`. | `frontmostApplication.bundleIdentifier == "com.microsoft.Powerpoint"`, plus an AXObserver on focus changes. |
| **Stale COM state** | Each focus or foreground event invalidates the snapshot to Unknown (pass) until a refresh lands. A slideshow is also checked live where it matters (the update notice uses window classes `screenClass`/`PodiumParent`). | Mac reads mode by AppleScript **on every owned key, synchronously** (`modeRead`), so there is no stale cache, but there is per-key latency (unverified). |
| **Conflicts with native shortcuts** | On the canvas the helper takes Ctrl+] [, Ctrl+Shift+C/V/T, Ctrl+K, Ctrl+Space, Ctrl+2–7, Alt+2/3, Alt+.; off the canvas, native behavior stays. | Mac takes Cmd+] [, Cmd+Shift+] [, Cmd+Shift+C/V, Cmd+/, Option+.. Check for conflicts with macOS-wide shortcuts (for example Cmd+Shift+C in some apps; PowerPoint for Mac native bindings). **Unverified.** |
| **Repeated presses** | Auto-repeat of a consumed key is swallowed. `H±`/`V±` repeat with throttling. | `swallowedKeycode` swallows repeats; `repeatGate` caches the decision. No spacing repeat. |
| **Hook lifecycle** | `WH_KEYBOARD_LL` on a message-loop thread. Unhooked in `finally`. Exceptions pass the key. Windows can silently remove a slow LL hook (`LowLevelHooksTimeout`), and there is **no detection** of that. COM calls run on the hook thread. | The CGEventTap is re-enabled on `tapDisabledByTimeout`/`ByUserInput`, plus a 0.25 s health timer. Better than Windows here; keep it. |
| **Layouts** | Letters by virtual key; punctuation from the active layout (`KeyLayout.Build` via `ToUnicodeEx`); US punctuation kept when the layout types a letter there (Arabic). AltGr: synthetic Left Ctrl detection (`AltGrRule`). | Letters by ANSI keycode; `/ . + -` from the active layout (`UCKeyTranslate`); `\ ; [` by physical keycode. There is no AltGr equivalent on Mac. |
| **Injected keys** | `LLKHF_INJECTED` always passes, so the helper's own replay never loops. | Unverified. Check `CGEventSourceStateID` / event source to avoid acting on synthetic events. |
| **Multiple PowerPoint windows** | The gate uses the foreground pid; COM attaches via `GetActiveObject`; replay is pinned to the frame hwnd. Multiple PowerPoint processes are not specifically handled. | AppleScript targets "Microsoft PowerPoint" (single app instance on Mac). |
| **Protected View / read-only** | Not handled explicitly (unverified). | Unverified. |

## 5. What is Windows-specific (do not port)

- `WH_KEYBOARD_LL`, `GetAsyncKeyState`, `LLKHF_*`, `SetWinEventHook`, the message-only window, `SetTimer`, `PostMessage`.
- Window class names (`PPTFrameClass`, `paneClassDC`, `mdiClass`, `NetUIHWND`, …). Mac must keep its AX role and subrole checks.
- `ToUnicodeEx` and the AltGr synthetic-Ctrl rule.
- `SendInput` text replay. The Mac `T` uses synchronous dispatch instead; only adopt replay if a real Mac problem shows up.
- COM `GetActiveObject` and the disconnect HRESULTs.

## 6. Portable logic (reuse as a specification, not as code)

`src/PPTKKeys/windows/ShortcutGate.cs` has **no Win32 and no COM**. Its header says it was written so the matrix can be tested on macOS. Use it as the behavioral spec and its tests as the test list.

- `ShortcutGate.Decide`, `DecideIcon`, `DecideSaveSlides`: what is allowed where.
- `ShortcutGate.Classify`: pane and selection to surface.
- `ShortcutGate.Reduce`: the T text-hold state machine.
- `ShortcutSequence` (`Advance`, `Next`, `Match`, `Expired`): sequence semantics.
- `ShortcutRepeat`: spacing repeat throttling.
- `DispatchPacing`: grid command gap.
- Tests: `PPTKKeys.Tests/ShortcutGateTests.cs`, `SequenceTests.cs`, `RepeatTests.cs`, `DispatchPacingTests.cs`. Port the **cases** to Swift XCTest. The Mac helper has no tests today.
