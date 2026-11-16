# 01 — Windows feature inventory (PPTK 1.1.0)

Verified on branch `windows-development` at `49fe2f2` (2026-11-16), by reading source and git history. Nothing was run for this document. Exact keys are in `05-KEYBOARD-SHORTCUTS.md`.

## How to read the status column

| Status | Meaning |
|---|---|
| **Tested** | There is automated test evidence (C#/PowerShell test suites) and/or a recorded real-device check. |
| **Implemented, not fully validated** | The code is complete and comments describe a manual check (for example "checked by exporting…"), but there is no test suite and no recorded host test. |
| **Implemented, unverified** | The code exists. There is no evidence of testing beyond the commit message. |
| **Experimental / disabled** | Present but switched off or explicitly marked as unvalidated. |
| **Not implemented** | No code found on any branch. |

**Important:** there is **no automated VBA test suite**. All add-in features are validated only by the developer's manual use in PowerPoint (commit subjects such as `windows11 - Add Enclose Icons`, `finalize Smart Bar, Save Slides…`). The helper (C#) and the deployment (PowerShell) do have test suites. The commit messages have no bodies.

## What the shipped add-in contains

`src/PPTK-WINDOWS-FINAL.pptm` (built into `PPTK.ppam`) imports only these components:

- **Modules:** `ModuleObjects`, `ModuleObjectsAlignAndDistribute`, `ModuleObjectsSizeAndPosition`, `ModulePPTKBrand`, `ModulePPTKFixColors`, `ModulePPTKFlip`, `ModulePPTKGrid`, `ModulePPTKReplaceColors`, `ModulePPTKRibbon`, `ModulePPTKShortcuts`, `ModulePPTKSlideLibrary`, `ModulePPTKSmartAlign`, `ModulePPTKSmartBar`, `ModulePPTKTakeaway`.
- **Classes:** `PPTKColorSwatch`, `PPTKMatrixCell`, `PPTKSmartBarRow`.
- **Forms:** `PPTKCentralForm`, `PPTKColorsForm`, `PPTKMatrixForm`, `PPTKSlideLibraryForm`, `PPTKSmartBarForm`.

Every other `.bas/.cls/.frm` under `src/` is inherited Instrumenta source that is **not** shipped on Windows. `src/PPTK/Ribbon/UnsupportedControls.txt` lists the Instrumenta ribbon controls left out. The Windows ribbon is `src/PPTK/Ribbon/CustomUI.xml`. `src/CustomUI/CustomUI.xml` is the inherited Instrumenta ribbon, which Windows does not use.

The VBA has **no `Declare` statements** (Defender ASR rule `92E97FA1` blocks Win32 calls from Office macros on company devices) and **no `#If Mac`** in shipped modules. `ModulePPTKBrand.bas` says "Do not import this module into the Mac workbook".

## A. Keyboard-driven features (PPTKKeys helper → `PPTK_DispatchShortcut`)

| Feature | Purpose / behavior | Entry points | Source | Status |
|---|---|---|---|---|
| **Quick insert shapes** | Insert a textbox, rectangle, rounded rectangle, oval, chevron, pentagon, triangle or parallelogram, centered. Repeated inserts are nudged 12 pt diagonally (tag `PPTK_SHORTCUT`). | T R U O C P G M; Insert group / Shapes menu | `ModulePPTKShortcuts.bas`: `PPTK_InsertShape`, `PPTK_InsertTextbox`, `PPTK_Insert*`, `PPTK_PlaceOrigin` | Tested (gate unit tests + daily use) |
| **T then type** | Characters typed right after T go into the new textbox. | T | helper `ShortcutGate.Reduce`, `ReplayGuard`, `Program.ReplayQueuedText` | Tested (unit) |
| **New slide, same layout** | Insert a slide after the current one with the same layout. | N; Insert group | `PPTK_InsertNewSlide` | Tested |
| **Insert Icon** | Opens the Windows Insert Icons dialog (`ExecuteMso "IconInsertFromFile"`; IconsPane is not used on Windows). | I; Insert group | `PPTK_InsertIcon`; helper `PowerPointSession.BeginInsertIcon` (separate STA thread) | Tested |
| **Enclose Icons** | 1st press: an oval of 1.5× the icon's bounding box behind the icon. Next presses toggle an Accent 1 fill. State is kept in tags `PPTK_ENCLOSURE_FOR`, `PPTK_ENCLOSURE_STATE`. | E; Insert group | `PPTK_EncloseIcons`, `PPTK_EncloseOne`, `PPTK_CreateEnclosure`, `PPTK_FindEnclosure` | Implemented, unverified (no undo entry; icons inside groups are ignored) |
| **Align / same size / distribute / spacing** | Sequences A·L/R/T/B/M/C, S·W/H/S, D·H/V, H·±, V·± (hold to repeat). | Align, Space, Shape groups | inherited `ModuleObjectsAlignAndDistribute.bas`, `ModuleObjectsSizeAndPosition.bas` (`SetRealSize` added for rotated shapes) | Tested (sequence unit tests + use) |
| **Match corner roundness** | Make rounded rectangles share the same visible corner radius (`Adjustments(1) × min(W,H)`). | A·A; Shape group | `PPTK_MatchCornerRoundness` | Implemented, not fully validated ("checked by exporting" comment) |
| **Smart Align** | Process (one flow with equal gaps), Matrix (grid picker 1–12), Central (circle around a centroid). | A·P, A·X, A·Z; Align group | `ModulePPTKSmartAlign.bas` + `PPTKMatrixForm`, `PPTKCentralForm` | Implemented, unverified (no undo entry) |
| **Unify shape type** | Give every selected shape the first shape's AutoShape type, keeping position, size and rotation. | X·T; Shape group | `PPTK_UnifyShapeType`, `PPTK_AssignAutoShapeType` | Implemented, unverified |
| **Merge text** | Merge the text of the selected shapes into the first one, then delete the sources. | . then T; Text group | `PPTK_MergeText` | Implemented, unverified |
| **Swap position / swap text** | Swap two shapes' positions (z-order kept, `PPTK_SwapZOrder`) / swap their text. | `/`, Ctrl+/ | `ObjectsSwapPosition`, `PPTK_SwapText` | Tested (gate) / unverified (VBA) |
| **Copy / paste formatting** | `ShapeRange.PickUp` / `Apply`. | Ctrl+Shift+C/V; Design group | `PPTK_CopyFormatting`, `PPTK_PasteFormatting` | Tested |
| **Paste content only** | Replace the selected text, or one shape's whole text, with plain clipboard text, keeping the formatting. | Ctrl+Shift+T (also while editing text); Text group | `PPTK_PasteContentOnly`, `PPTK_ClipboardText` (MSForms DataObject) | Implemented, unverified |
| **Remove margins (zero margins)** | Zero the text margins, then shrink the shape or table rows/columns to fit the text. | Alt+. (also while editing text); Text group | `PPTK_ZeroMargins` | Implemented, unverified |
| **Font size / z-order** | `ExecuteMso FontSizeIncrease/Decrease`, `ObjectBringForward/SendBackward`. | Ctrl+] [, Ctrl+Shift+] [ | `PPTK_ExecuteMso` | Tested |
| **Grid rows/columns (keep bounds)** | Add or delete the right column / bottom row of a grid of shapes or a native table, keeping the outer rectangle. Deletes keep hidden recovery copies. | Ctrl+2/3, Alt+2/3; Rows/TK group | `ModulePPTKGrid.bas` (`PPTK_Recognize`, `PPTK_AddRightColumn`, …) | Implemented, unverified. The 100 ms dispatch pacing exists because of a real crash (see 06). |
| **Grid expand/shrink (legacy geometry)** | The same, growing or shrinking outward. | Ribbon "legacy" menu; Smart Bar | `PPTK_*Expand`, `PPTK_*Shrink` | Implemented, unverified |
| **TK placeholders** | Insert one editable "TK" text box beside the selection: right, bottom, top or left. | Ctrl+4/5/6/7; Rows/TK group | `ModulePPTKTakeaway.bas` (`PPTK_AddTK`) | Implemented, unverified |
| **Fix Slide Colors** | Map hard-coded RGB colors on the current slide to theme colors (Light/Dark 1/2, nearest accent), forcing a minimum 3:1 text contrast. | Q; Design group | `ModulePPTKFixColors.bas` (`PPTK_FixSlideColors`); specification `ThemeHarmonize.cs` (C#, tested) | Implemented, not fully validated (undo explicitly unverified in a comment) |
| **Smart Bar** | Command palette with 60 commands, context suggestions and ranked search; keeps the selection. | Ctrl+K, Ctrl+Space; Tools group | `ModulePPTKSmartBar.bas`, `PPTKSmartBarForm`, `PPTKSmartBarRow` | Implemented ("finalize" commit); no test record |
| **Save selected slides** | Save the thumbnail-pane selection as a new .pptx (Unicode/Arabic paths work). | `\` (thumbnail pane only); Reuse group | `PPTK_SaveSelectedSlides`, `PPTK_KeepSlideIndexes` | Implemented ("finalize" commit); unverified |

## B. Ribbon-only features (no keyboard shortcut)

| Feature | Behavior | Ribbon | Source | Status |
|---|---|---|---|---|
| **Find & Replace Colors** | Scan fill, outline and font colors over all slides, the current slide or a range; replace by exact RGB. Takes a backup to `Documents\PPTK Color Backups` first. One undo entry. | Design → Find & Replace Colors; Smart Bar | `ModulePPTKReplaceColors.bas`, `PPTKColorsForm`, `PPTKColorSwatch` | Implemented, not fully validated (comment: checked against rendered pixels) |
| **Flip slide / flip presentation** | Mirror layouts for right-to-left: positions, rotation, `Shape.Flip` for AutoShapes, paragraph alignment, table direction, SmartArt reverse, re-glued connectors. Running it twice restores the original. Whole-deck: confirmation + backup. | Tools → Flip; Smart Bar | `ModulePPTKFlip.bas` | Implemented, not fully validated (comment: export check with Arabic and English) |
| **Email selected slides** | Classic Outlook draft via COM; falls back to `Documents\PPTK Email`. | Reuse → Email Slides | `PPTK_EmailSelectedSlides` | Implemented, unverified (uses ANSI `Dir$` for paths) |
| **Slide Library** | Save the current slide(s) as PPTX + JPG preview into a library folder (default `Documents\PPTK Slide Library`; setting `PPTK\SlideLibrary\Folder`); browse and insert. | Reuse → Library menu | `ModulePPTKSlideLibrary.bas`, `PPTKSlideLibraryForm` | Implemented, not fully validated (`src/Forms/PPTKSlideLibraryForm.log` shows form import errors; the form's `.frx` is not in the repo) |
| **Inherited align/space/size tools** | Stretch, ResizeAndSpaceEven, RemoveSpacing, ArrangeShapes, SizeTo*, swap centered. | Align, Space, Shape groups | inherited modules | Inherited from Instrumenta; unverified on Windows |

## C. Helper and platform features (Windows)

| Feature | Source | Status |
|---|---|---|
| Keyboard hook, context gate, sequences, chords, layouts, AltGr | `src/PPTKKeys/windows/*` (see 02, 05) | Tested (pure logic: ~58 unit tests); real hook validated in manual use |
| Pending add-in update applied when PowerPoint closes | `PendingAddinUpdate.cs`, `PendingUpdateWatcher.cs` | Tested (unit + script contract) and a real check with test names |
| Update notice (no focus steal, 10 s, once per version) | `UpdateNotice.cs`, `NotificationWindow.cs`, `PowerPointWindows.cs` | Tested + real-device check (slideshow fix verified live) |
| Helper handoff (no gap, never two hooks) | `HelperHandoff.cs`, `HelperWiring.cs` | Tested (real processes, fake hook) + a real-hook check with test names |
| Forced termination of a stuck new helper | `HelperHandoff.RunOld`, `DeploymentFlags.ForcedRecoveryEnabled = false` | **Disabled**, not validated |
| Per-user installer, uninstaller, detection, rollback, Intune package, direct ZIP | `src/Deploy/*`, worktree branch `direct-installer` | Tested in isolation; Intune pilot not done (see 07) |
| Universal x64/ARM64 private runtime | `install.ps1`, `deploy-env.ps1`, `build-package.ps1` | Tested on x64. ARM64: only files and signatures verified. |

## D. Dead, missing or not implemented

| Item | Finding |
|---|---|
| **Linked Groups / synchronization** | **Not implemented.** No code, commit or document on any branch (`git log --all`, source grep). |
| **Shape replacement ("Replace Shape")** | **Not implemented** as a shape tool. The only match is the inherited `ReplaceShapeColor` (`src/Modules/ModuleColorScanner.bas`, not shipped). The nearest PPTK feature is *Unify Shape Type* (X·T). |
| Align Shape Angle (`windows8`) | Superseded. `ObjectsAlignShapeAngle` (`ModuleObjectsAlignAndDistribute.bas`) remains but is unreachable; A·A is now Match Corner Roundness. The icon `pptk-align-angle.png` is orphaned. |
| Inherited Instrumenta tools (stylesheets, crop, sticky notes, Harvey balls, tables, Slide Grader, mail merge, agenda, script, …) | In source, **not shipped on Windows** (`UnsupportedControls.txt`). |
| `src/PPTKKeys/windows/README.md` | **Stale:** describes only 9 letters and says the helper was "compiled on macOS" and never run on Windows. |
| `HelperWiringTests.Helper_informational_arguments_behave_as_before` | Asserts `--version == "1.0.0"`, but `PPTKKeys.csproj` is now `1.1.0`. Likely failing after the 1.1.0 bump (not run). |

## E. State persistence (Windows)

- **Registry (VBA `SaveSetting`):** `PPTK\SlideLibrary\Folder`. Inherited modules read `Instrumenta\AlignDistributeSize\*` and fall back to defaults; the settings UI is not shipped.
- **Shape tags:** `PPTK_SHORTCUT`, `PPTK_ENCLOSURE_FOR`, `PPTK_ENCLOSURE_STATE`. No document-level tags.
- **Files:**
  - `Documents\PPTK Color Backups`, `PPTK Flip Backups`, `PPTK Slide Library`, `PPTK Email`;
  - `%TEMP%\PPTK-slides-*.pptx`.
- **Deployment state:** `%LOCALAPPDATA%\PPTK\state.json` and related files (see 07).
