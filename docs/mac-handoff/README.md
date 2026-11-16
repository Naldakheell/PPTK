# PPTK — Windows → Mac engineering handoff

Prepared from the actual repository on 2026-11-16:

- **Windows:** `windows-development` at `49fe2f2`, PPTK 1.1.0.
- **Mac:** `origin/mac-development` at `4503753 mac1`. The local ref is from 2026-11-04; fetch before relying on it.

This is documentation only. Nothing in Windows or Mac code was changed, built, or run for it.

## Executive summary

- **Architecture is already shared.** Both platforms use a native keyboard helper that decides context and sends a **token** to one VBA entry macro, `PPTK_DispatchShortcut(token)`.
  - Windows: C# `WH_KEYBOARD_LL` + COM `Application.Run`.
  - Mac: Swift CGEventTap + AppleScript `run VB macro`.
  - The Mac project should be continued, not rewritten.
- **Mac already has the core set:**
  - inserts T R U O C P G M N I;
  - align / same-size / distribute / spacing sequences;
  - swap, merge, copy/paste formatting, zero margins;
  - font size and z-order chords;
  - save slides; plus Mac-only `;` (slide library) and `[` (email slides).
- **The Windows-only gap is mostly portable VBA:**
  - Grid rows/columns, TK placeholders, Smart Align;
  - Match Corner Roundness, Unify Shape Type, Enclose Icons;
  - Fix Slide Colors, Flip, Find & Replace Colors;
  - Paste Content Only, Smart Bar.

  Most of it is object-model code. The forms (Smart Bar, Colors, Matrix, Central) are the main Mac risk.
- **Token names differ** between the platforms for the same features (`CMD /` vs `CTRL /`, `FONT +` vs `CTRL ]`, …). Unify them in VBA first.
- **Windows-only by nature:**
  - the Win32 hook, focus probe and layout code;
  - COM;
  - the update notice, pending updates and handoff;
  - the entire installer and Intune stack (PowerShell, Task Scheduler, registry).

  The concepts are reusable for a future Mac distribution. The code is not.
- **Not implemented anywhere:** Linked Groups / synchronization and Replace Shape. No code, commit or doc exists on either branch.
- **Biggest lessons:**
  - Company security controls decide the architecture: ASR blocked an unsigned exe and VBA `Declare`s. Expect TCC, Gatekeeper and MDM to matter equally on Mac.
  - Fail open for keys.
  - Never trust cached state for timing.
  - Test "T then type".
  - Verify the shipped artifact.

## Reading order

| # | Document | Read it for |
|---|---|---|
| 1 | `01-WINDOWS-FEATURE-INVENTORY.md` | What Windows has, with status and sources |
| 2 | `04-FEATURE-PARITY-MATRIX.md` | What Mac already has and the gaps |
| 3 | `05-KEYBOARD-SHORTCUTS.md` | Exact keys, gates, edge cases, token differences |
| 4 | `02-WINDOWS-ARCHITECTURE.md` | How and why Windows is built this way |
| 5 | `03-REUSABLE-COMPONENTS.md` | What to port, adapt, or leave behind |
| 6 | `06-BUGS-AND-LESSONS.md` | Mistakes not to repeat on Mac |
| 7 | `08-MAC-CONTINUATION-ROADMAP.md` | What to do next, in order |
| 8 | `07-RELEASE-AND-DEPLOYMENT.md` | Windows packaging and its Mac-relevant concepts (later) |

## Before implementing anything, the Mac session must

1. Fetch and check out `mac-development` on a Mac. Do not merge `windows-development` into it.
2. Build and run the current Swift helper and record which shortcuts work (Phase 0 of 08).
3. Export the VBA from `src/PPTK DEV.pptm` and compare it with the `.bas` files on the Mac branch.
4. Confirm the TCC grants and how they are attributed (Terminal vs binary).
5. Measure per-key AppleScript latency and fast "T then type".

## Uncertainties (explicitly unverified)

- The Mac branch tip and its runtime behavior (nothing was run on a Mac).
- Mac `.pptm` contents versus the `.bas` files.
- Most Windows VBA features have no automated tests. Their status comes from code comments and commit subjects (see 01).
- Native PowerPoint shortcut conflicts, and the Alt KeyTips effect after swallowed Alt chords on Windows.
- Windows ARM64 runtime behavior and Intune pilot behavior.

## Known stale items in the repository (not changed here)

- `src/PPTKKeys/windows/README.md` describes only 9 letters and is outdated.
- `src/PPTKKeys/mac/` in the Windows checkout is the old 588-line helper. The current Mac helper is on `origin/mac-development`.
- `HelperWiringTests` expects `--version` 1.0.0. The project is 1.1.0.
- `ObjectsAlignShapeAngle` is unreachable dead code. Its icon `pptk-align-angle.png` is orphaned.
