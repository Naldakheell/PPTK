# 08 — Recommended continuation roadmap for the existing Mac project

**Principles, in priority order:**

1. Preserve what works on Mac today.
2. Close the most valuable gaps.
3. Reuse proven algorithms.
4. Adapt the platform layer.
5. Keep the user experience consistent.

**No rewrite.** Employee-feedback ideas that are not in code are not listed as Windows features.

## Phase 0 — Establish the real Mac baseline (before any code)

1. **Get the current Mac branch:**
   - `git fetch` and check `origin/mac-development`. The local ref is from 2026-11-04 and may be stale.
   - Check out the Mac branch on the Mac. Do not merge the Windows branch into it.
2. **Build and run the helper exactly as it is** (`swift build` / `swift run` in `src/PPTKKeys/mac`; there are no documented build steps — record them).
   - Grant Accessibility, Input Monitoring and Automation.
   - Verify every shortcut in the Mac table (05 §2 and the Mac section of 04) on a real Mac with PowerPoint. Record what works.
3. **Export the VBA from the Mac `PPTK DEV.pptm`** and diff it against `src/Modules/ModulePPTKShortcuts.bas` and the other `.bas` files on the Mac branch. This tells you whether the `.bas` files are the truth.
4. **Measure the per-key latency** of the synchronous `modeRead()` AppleScript on a typical deck.
5. **Check "T then type"** at fast typing speed.

**Deliverable:** a baseline test record. Without it, "parity" cannot be measured.

## Phase 1 — Align the shared contract (small, high value)

1. **One dispatcher accepting both token spellings** (`CMD /` and `CTRL /`, `FONT +` and `CTRL ]`, …).
   - Do this in VBA only.
   - Do not change the helper's keys.
   - See 05 §1.
2. **Decide the Mac shortcut set explicitly.** Keep the Mac-only `;` and `[`, or move them to the ribbon as on Windows. Record the decision; do not drop them silently.
3. **Reconcile `SetRealSize`:** compare the two implementations on rotated shapes and keep one.
4. **Port the gate test cases** from `ShortcutGateTests.cs`, `SequenceTests.cs` and `RepeatTests.cs` into XCTest against the Swift gate functions. Refactor the Swift gate into testable pure functions *only as far as needed* to test them.

## Phase 2 — Port the portable VBA features (in this order)

Each item is VBA with little or no Windows dependency (see 03 §A).

| Order | Feature | Steps |
|---|---|---|
| 1 | Match corner roundness (A·A), Unify shape type (X·T) | Copy, test on Mac, add tokens to the dispatcher and the helper sequences (`A A`, `X T` with an `X` prefix). |
| 2 | Grid rows/columns + pacing (Ctrl/Alt+2/3 → Cmd/Option+2/3?) | Pick Mac keys without conflicts. Test the hidden recovery copies. |
| 3 | TK placeholders | Same key-mapping decision as the grid. |
| 4 | Enclose icons (E) | Verify icon detection on Mac. |
| 5 | Smart Align P/X/Z | Geometry first; then the Matrix and Central forms on Mac. |
| 6 | Fix Slide Colors (Q) | Validate against `ThemeHarmonizeTests` cases. Check undo. |
| 7 | Flip slide/deck, Find & Replace Colors | Adapt the backup folder paths. Forms for Replace Colors. |
| 8 | Paste content only | Needs a Mac clipboard path; MSForms DataObject is unverified on Mac. |
| 9 | Smart Bar | Biggest UI risk (MSForms on Mac). Consider last. Check Cmd+K conflicts. |

For each feature: port → manual test matrix on a real Mac → add to the Mac ribbon → document.

## Phase 3 — Helper robustness (adapt Windows lessons)

1. **Swallow auto-repeat for every one-shot command** (already partly done). Add H±/V± hold-to-repeat with the Windows throttle (50 ms, stop on release) if wanted.
2. **Invalid second key re-processed as a new key; Esc swallowed** (Windows `ShortcutSequence.Next`). Align with the Mac prefix logic.
3. **Ignore synthetic events** (Windows ignores `LLKHF_INJECTED`). Check the event source on Mac.
4. **Single instance** (a lock file or named semaphore), so two helpers never run.
5. **A log file** (`~/Library/Logs/PPTK/`), with no key contents.

## Phase 4 — Mac distribution (later; do not start now)

1. App bundle + Info.plist (`NSAppleEventsUsageDescription`), codesign + notarization. These are required for stable TCC grants.
2. A LaunchAgent for auto-start.
3. Per-user `.pkg` or managed deployment, reusing the manifest + hash + journal concepts (07 §3).
4. Pending update on PowerPoint close + a non-activating notice: only after packaging exists.

## What not to do

- Do not replace the Mac CGEventTap / AX / AppleScript layer with ports of Win32 code.
- Do not remove Mac-only features (the Instrumenta tools, `;`, `[`) to match Windows.
- Do not copy `ModulePPTKBrand.bas` into the Mac workbook (its own header forbids it).
- Do not merge `windows-development` into `mac-development` wholesale. Port feature by feature.
- Do not count a Windows feature as "Mac ✔" until it was run in PowerPoint for Mac.
