# 06 — Bugs, root causes, fixes and lessons

**Sources:**

- code comments;
- `src/Deploy/*.md`;
- the PPTKKeys tests;
- the project memory note on PPTM builds;
- the Windows development sessions that built deployment and notification (this author).

Commit messages are subject-only and give no detail.

**Labels:**

- **Verified** — reproduced and fixed with a recorded check.
- **From code** — the reason is stated in a code comment.

## 1. Bugs and fixes

| # | Bug | Root cause | Fix | Regression scenario | Evidence |
|---|---|---|---|---|---|
| 1 | Unsigned `PPTKKeys.exe` blocked | Defender ASR rule `01443614` (low-prevalence executables) | Ship no apphost. Run `dotnet.exe exec PPTKKeys.dll`, and in 1.1.0 a Microsoft-signed private runtime. | Any new unsigned binary in the release | Verified (9 ASR events naming `PPTKKeys.exe`) |
| 2 | PPTM saves silently failed; the add-in went stale | ASR rule `92E97FA1` blocks Office macros with `Declare … Lib "user32"` | All Win32 `Declare`s removed from the shipped VBA. The Smart Bar uses Hide/Show and a custom scrollbar instead. | Adding a `Declare` | Verified (project memory, ASR log) |
| 3 | The re-shown Smart Bar lost typing focus | MSForms modeless form reuse | Always unload and recreate the form (`PPTK_ShowSmartBar`) | Open / close / reopen the Smart Bar quickly | From code |
| 4 | Office crash in `mso20win32client.dll` | Ctrl+2 run 0 ms after Alt+2 (grid add right after delete) | `DispatchPacing.GapMs = 100` after grid tokens | Rapid grid key presses | From code |
| 5 | Save Slides reported failure on Arabic paths | `Dir$` is ANSI | `FileSystemObject` for existence checks | Arabic file names. **Still open in Email Slides** (`PPTK_PathExists` uses `Dir$`). | From code |
| 6 | The update notice stayed hidden 30 s after a slideshow ended | The cached COM snapshot refreshes only on PowerPoint events | Live check for visible `screenClass` / `PodiumParent` windows (`PowerPointWindows.HasVisibleSlideshow`) | Start a slideshow, press Esc, expect the notice within ~5 s with no click | Verified on a real device |
| 7 | Update-notice ANSI window problem | `DefWindowProc` declared ANSI | Unicode `DefWindowProc` | Notice text rendering | Verified (tests) |
| 8 | The notice could show twice | A race between the scheduler and a stale inspection | `Appeared` recorded on the message thread before the next tick | Fast ticks around display | Verified (tests) |
| 9 | The undo could not delete the new version folder | The new helper was already running from it (the DLL was mapped) | The undo asks the helper to close first; version folders are removed best-effort | A failed upgrade after the helper restart | Found by code review; fixed + tests |
| 10 | Install / uninstall continued after "refused" | PowerShell `exit` in a dot-sourced file leaves only that file | `[Environment]::Exit` | Unsafe root / SYSTEM context must stop the script | Verified (tests) |
| 11 | PowerShell function results polluted | Functions returning bool into the pipeline; `@()` flattening single pairs | Explicit `[void]`; an explicit list for registry values | A single-value registry key | Verified (live tests) |
| 12 | Helper discovery found nothing | `$null` passed as `""` to Win32 `FindWindowEx`; the helper window title is registered ANSI ("P") | Pass `[NullString]::Value`; match by class `PPTKKeys.Message.<pid>` | — | Verified |
| 13 | Wrong architecture under emulation | x64/x86 PowerShell on ARM64 reports the emulated architecture | Native architecture from `Win32_Processor` / the registry | Install from a 32-bit or emulated host | Verified (tests) |
| 14 | Relative paths wrong | Short 8.3 paths (`NOUFAL~1`) versus long paths | Normalize to the long form before computing relative paths | — | Verified |
| 15 | `-Include *.exe` returned text files | `Get-ChildItem -LiteralPath -Include` is unreliable | Filter by `Extension` | Signature check of the runtime | Verified |
| 16 | Live Task Scheduler / registry tests looked like they "failed restore" | Registry values are restored in a different order | Order-independent comparison | — | Verified |
| 17 | Distribute / align on grouped children misbehaved | The selection was read from `ShapeRange` instead of `ChildShapeRange` | Child-range first (inherited modules) | Select shapes inside a group, then align | From code |
| 18 | Same size broke rotated shapes | Width / height of the unrotated box | `SetRealSize` keeps the visible size and center; warns when unreachable | A 45° rotated rectangle | From code (independent fix on Mac too) |

## 2. Lessons for the Mac session (most important first)

1. **Corporate security decides what can ship.**
   - On Windows, ASR rules killed two designs: the apphost exe and VBA Declares.
   - On Mac, expect the same from TCC (Accessibility, Input Monitoring, Automation of PowerPoint), Gatekeeper/notarization, and MDM profiles.
   - **Test on a managed company Mac first**, before adding features.
2. **Keep the token contract.** One entry macro, string tokens, VBA does the work. Both platforms independently converged on it.
3. **Fail open for keys.** If context is unknown, pass the key. A swallowed keystroke while the user types is the worst bug class.
4. **Never trust cached state for timing-sensitive decisions** (bug 6). The Mac helper reads PowerPoint mode per key via AppleScript. That is correct, but measure the latency.
5. **"T then type" needs a design** (see 05 §4). Windows needed a state machine plus a guarded replay; Mac uses synchronous dispatch. Test typing speed explicitly.
6. **Every one-shot command must swallow its own auto-repeat**, otherwise one long press inserts five shapes.
7. **The hook/tap can die.**
   - Windows has no detection of a silently removed low-level hook (open risk).
   - Mac already re-enables the tap on timeout. Keep it.
8. **Unicode paths:** Arabic file names broke ANSI APIs on Windows (bug 5). On Mac, test Arabic paths in `MacScript`/AppleScript calls.
9. **Arabic keyboard layout:**
   - Letters work positionally.
   - Punctuation must come from the layout, except where the layout types a letter there; then the US meaning is kept.
   - The Mac helper already handles `\ ; [` physically. Keep that.
10. **Undo grouping is inconsistent on Windows** (see 02). When porting a feature, decide undo behavior explicitly and test it in PowerPoint for Mac.
11. **Forms are a risk area.** MSForms UserForms differ on Mac. The Windows Smart Bar, Colors, Matrix, Central and Slide Library forms all need Mac validation before they count as ported.
12. **Verify the shipped artifact, not only the source.** The Windows build compares the PPTM's VBA with the source modules and fails when stale.
    - On Mac, `PPTK DEV.pptm` contents versus the `.bas` files are unverified today.

## 3. Remaining platform risks (Windows, for awareness)

- **Silent hook removal:** a low-level hook can be removed by Windows when it is slow, and there is no recovery (no code detects it).
- **Forced recovery during handoff** is disabled; it was never validated.
- **Intune behavior not piloted:** the 1618 retry window, user-context detection, ASR on pilot devices.
- **ARM64 Windows:** never executed on a real device.
- **Stale test:** `HelperWiringTests` expects `--version` 1.0.0 (now 1.1.0).
- **Stale README:** `src/PPTKKeys/windows/README.md`.
- **No VBA test suite:** all add-in features rely on manual validation.
