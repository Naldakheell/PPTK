# 04 — Preliminary Windows → Mac feature parity matrix

**Mac evidence:** `origin/mac-development` at `4503753 mac1` (2026-11-04). Read with `git show`, never checked out or run.

**Freshness of that ref:**

- The local remote-tracking ref last moved on 2026-11-04.
- A later fetch probably saw no new commit.
- **Whether this is still the remote tip is unverified.** Fetch first.

**What the Mac branch contains (beyond the shared base `14f531b`):**

- the Swift helper `src/PPTKKeys/mac/Sources/PPTKKeys/main.swift` (1253 lines);
- `src/Modules/ModulePPTKShortcuts.bas` (320 lines);
- `src/Modules/ModulePPTKRibbon.bas`;
- small changes to three inherited modules;
- `src/PPTK DEV.pptm`.

**Two things are unverified:**

- The VBA *inside* `PPTK DEV.pptm` was not decompiled, so whether it matches the exported `.bas` files is **unverified**.
- **Nothing was run on a Mac** for this document.

> The `src/PPTKKeys/mac/` copy in the Windows checkout is the **old** 588-line version from `14f531b` (9 letters only). Do not use it; the Mac branch is newer.

## Legend

| Class | Meaning |
|---|---|
| **Mac ✔** | Already implemented on Mac (code evidence). |
| **Mac ~** | On Mac but needs improvement or alignment. |
| **Win-only** | Windows-specific by nature; not to be ported as-is. |
| **Investigate** | A Windows feature that is portable in principle; needs a macOS-specific investigation. |
| **Unverifiable** | Cannot be confirmed from this checkout. |

## Keyboard features

| Feature | Windows | Mac evidence | Class |
|---|---|---|---|
| Insert T R U O C P G M | ✔ | `main.swift` `directTokens`; Mac `PPTK_Insert*` | **Mac ✔** |
| T then typing | text hold + replay | synchronous dispatch of `T` (line 522) | **Mac ~** (fast-typing behavior unverified) |
| N new slide | ✔ | ✔ | **Mac ✔** |
| I insert icon | Insert Icons dialog | `ExecuteMso "IconsPane"` | **Mac ✔** (platform-specific command) |
| Align A·L/R/T/B/M/C | ✔ | ✔ | **Mac ✔** |
| Same size S·W/H/S | ✔ (`SetRealSize`) | ✔ (its own `SetRealSize`) | **Mac ~** (two independent implementations) |
| Distribute D·H/V | ✔ | ✔ | **Mac ✔** |
| Spacing H±, V± | ✔ + hold to repeat | ✔, no repeat | **Mac ~** |
| Swap position `/`, swap text | ✔ | ✔ (Cmd+/) | **Mac ✔** |
| Merge text `. T` | `PPTK_MergeText` | Instrumenta `ObjectsTextMerge` | **Mac ✔** (different implementation) |
| Copy / paste formatting | Ctrl+Shift+C/V | Cmd+Shift+C/V | **Mac ✔** |
| Zero margins | Alt+. (+ fit to text) | Option+. via `ObjectsMarginsToZero` | **Mac ~** (no fit-to-text) |
| Font size, z-order chords | Ctrl+] [, Ctrl+Shift+] [ | Cmd+] [, Cmd+Shift+] [ (font also while editing text) | **Mac ✔** (token names differ) |
| Save selected slides `\` | thumbnail pane | slide selection; Mac cancel fix in `ModuleExport.SaveSelectedSlides` | **Mac ✔** |
| Slide library `;` | not a shortcut (ribbon) | `ShowSlideLibrary` (Instrumenta) | Mac-only shortcut; keep |
| Email slides `[` | ribbon only | `EmailSelectedSlides` via the AppleScriptTask plugin | Mac-only shortcut; keep |
| Enclose icons `E` | ✔ | – | **Investigate** (icon type detection on Mac) |
| Fix slide colors `Q` | ✔ | – | **Investigate** |
| Match corner roundness A·A | ✔ | – | **Investigate** (likely direct VBA port) |
| Smart Align A·P/X/Z | ✔ (+2 forms) | – | **Investigate** (forms on Mac) |
| Unify shape type X·T | ✔ | – | **Investigate** (likely direct port) |
| Paste content only Ctrl+Shift+T | ✔ (MSForms DataObject) | – | **Investigate** (Mac clipboard access from VBA) |
| Grid add/delete Ctrl/Alt+2/3 | ✔ | – | **Investigate** (largest gain) |
| TK placeholders Ctrl+4–7 | ✔ | – | **Investigate** |
| Smart Bar Ctrl+K / Ctrl+Space | ✔ (MSForms palette) | – | **Investigate** (MSForms limits on Mac; Cmd+K conflicts) |

## Ribbon-only features

| Feature | Windows | Mac | Class |
|---|---|---|---|
| PPTK ribbon tab | full (`src/PPTK/Ribbon/CustomUI.xml`) | Instrumenta ribbon + 2 PPTK buttons (T, N) in `PPTK DEV.pptm` | **Mac ~** |
| Find & Replace Colors | ✔ | – | **Investigate** (form) |
| Flip slide / presentation | ✔ | – | **Investigate** |
| Slide Library (PPTK version) | ✔ | Instrumenta slide library | **Mac ✔** (different implementation; keep Mac's) |
| Email slides | Outlook COM | AppleScriptTask plugin | **Mac ✔** (platform-specific) |
| Inherited Instrumenta tools (stylesheets, crop, tables, …) | not shipped on Windows | available on Mac via the Instrumenta ribbon | **Mac ✔** (Mac has more here; do not remove) |

## Helper, platform and deployment

| Feature | Windows | Mac | Class |
|---|---|---|---|
| Global key capture | `WH_KEYBOARD_LL` | CGEventTap with re-enable + health timer | **Mac ✔** |
| Context gate | Win32 classes + COM snapshot | AX roles + per-key AppleScript `selection type` | **Mac ✔** (latency unverified) |
| Layout handling | `ToUnicodeEx`, Arabic rule, AltGr | `UCKeyTranslate`, physical keys for `\ ; [` | **Mac ✔** |
| Single instance | named mutex | – | **Investigate** |
| Auto start | Task Scheduler at logon | manual CLI only | **Investigate** (LaunchAgent) |
| Permissions | – | Accessibility, Input Monitoring, Automation (TCC) | Mac-only concern; app bundle + signing needed for stable grants |
| Logging | `%LOCALAPPDATA%\PPTK\logs` (production) | stderr only | **Investigate** |
| Pending update + notice + handoff | ✔ | – | **Investigate** (only after Mac packaging exists) |
| Installer / Intune / ZIP | ✔ | Instrumenta `.pkg` scaffolding only (`src/Installers/pkg/`) | **Win-only** code; concept reusable later |
| Tests | ~58 pure gate tests; deployment suites | none for the Swift helper | **Investigate** (port the gate test cases to XCTest) |

## Not implemented anywhere

| Feature | Note |
|---|---|
| Linked Groups / synchronization | No code on either branch |
| Replace Shape | No code on either branch (only the inherited `ReplaceShapeColor`) |
