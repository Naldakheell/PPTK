# PPTK transaction journal (`LocalAppData\PPTK\transaction.json`)

One journal format is shared by every PPTK operation that changes an installation:
`install.ps1` (install, rollback), `uninstall.ps1`, and the PPTKKeys `PendingAddinUpdate`
engine. Any of them must be able to undo an operation another one left interrupted.

## Rules

- Only one operation runs at a time: each holds `LocalAppData\PPTK\state.lock`, opened with
  `FileMode.OpenOrCreate`, `FileAccess.ReadWrite`, `FileShare.None`. Windows releases it when the
  process ends, even if it is killed.
- The journal is written (atomically: temporary file, then replace) **before** the first change,
  and deleted only after the operation is verified. Its presence means "interrupted".
- The next operation that takes the lock undoes the journal first, then starts its own work.
  The PPTKKeys engine undoes only its own journals (`operation = apply-pending`); for any other
  operation it does nothing and waits for the scripts.
- The undo restores everything recorded below, verifies the add-in hash, then deletes `txn\`
  and the journal. If the undo fails, the journal is kept and the failure is reported
  ("RESTORE FAILED"); nothing guesses.

## Fields (journalVersion 1)

| Field | Meaning |
|---|---|
| `journalVersion` | `1`. Missing means 1 (journals written before this field existed). Any other value: the undo refuses. |
| `tasksFormat` | Format of `tasksBefore`. Missing means `sim-json`. An undo refuses a format its environment does not use. |
| `operation` | `install`, `rollback`, `uninstall`, `cancel-pending`, `apply-pending`. |
| `target` | The version the operation was working on. |
| `startedUtc` | ISO 8601 time. |
| `stateBefore` | Full text of `state.json` before the operation, or `null` (no file: delete it on undo). |
| `tasksBefore` | PPTKKeys task(s) before the operation in `tasksFormat`, or `null` (no PPTKKeys task: on undo remove only PPTKKeys tasks, never other tasks). |
| `pptkKeyBefore` | `HKCU\Software\PPTK` values as a JSON object text, or `null` (no key: delete it on undo). |
| `addinKeyBefore` | `HKCU\Software\Microsoft\Office\16.0\PowerPoint\AddIns\PPTK` values as JSON object text, or `null`. |
| `addinBefore` | Full path of the copy of `AddIns\PPTK.ppam` in `txn\`, or `null` (no add-in: delete it on undo). |
| `addinBeforeSha` | SHA-256 (lowercase hex) of that copy. The undo checks it before and after restoring, and does not overwrite an add-in that already has this hash. |
| `createdVersionDir` | `versions\<v>` created by this operation (deleted on undo), or `null`. |
| `moved` | uninstall only: `[{from, to}]` folders moved into `txn\removed\`; moved back on undo. |

## tasksFormat

- `sim-json` — test roots (`-Root`): the full text of `<Root>\tasks.json`. Used by every script and
  test today.
- `task-xml` — the live environment (production, and `-LiveTestId` tests under `PPTK-Test-<id>`):
  `{"tasks":[{"path":"\\","name":"PPTKKeys Startup","xml":"<Task ...>"}]}`, containing only the main
  task (`PPTKKeys Startup`; the test task in live tests). Written by `install.ps1`/`uninstall.ps1`
  (`Export-ScheduledTask`, `src\Deploy\deploy-env.ps1`) and by the PPTKKeys engine (Task Scheduler
  COM `Xml`). Restored for that task name only: scripts with `Register-ScheduledTask -Xml -Force`,
  the engine with `RegisterTask(..., TASK_CREATE_OR_UPDATE, ..., TASK_LOGON_INTERACTIVE_TOKEN)`.

## Registry values

JSON numbers are `REG_DWORD`, JSON strings are `REG_SZ` (e.g. `{"Path":"PPTK.ppam","AutoLoad":1}`).
