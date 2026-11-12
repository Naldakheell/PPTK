# PPTKKeys Windows installer notes

This is the contract for a future per-user PPTK installer. It does not describe an MSI. The installer must reuse the existing PPTKKeys application and must not add a second keyboard helper.

The current machine is already configured this way. The notes below match that configuration.

## Install the helper

Copy the latest tested Release build into a stable per-user directory:

`%LOCALAPPDATA%\PPTK\PPTKKeys`

Required files, kept together:

- `PPTKKeys.dll`
- `PPTKKeys.deps.json`
- `PPTKKeys.runtimeconfig.json`

Do not install from `bin\Release\compile-check` or any other temporary compile folder. That folder is only a build output. The permanent location is the directory above.

Do not ship the apphost `PPTKKeys.exe` as the startup program. On this machine `dotnet.exe` is not under `C:\Program Files\dotnet`, and the apphost exits because it does not see the user-local runtime. Launch with the full path to `dotnet.exe`.

## Find the .NET runtime

The tested build targets `net10.0-windows`. Its `PPTKKeys.runtimeconfig.json` requires:

- framework `Microsoft.NETCore.App`
- minimum version `10.0.0`

The installer must not assume `C:\Program Files\dotnet\dotnet.exe`. Search in this order:

1. The `dotnet.exe` recorded from a previous successful run on this user.
2. `%LOCALAPPDATA%\Microsoft\dotnet\dotnet.exe`
3. `dotnet.exe` on the user `PATH`
4. `C:\Program Files\dotnet\dotnet.exe`

Confirm the runtime with that executable:

`dotnet.exe --list-runtimes`

Accept a roll-forward inside .NET 10. On this machine `Microsoft.NETCore.App 10.0.14` satisfies the `10.0.0` requirement. `Microsoft.WindowsDesktop.App` is installed here but is not listed in the current `runtimeconfig.json`. Read the file. Do not hard-code an extra framework.

If the required framework is missing, `dotnet exec` fails and prints that the framework was not found. The installer should show that message and stop. It should not change Defender, Attack Surface Reduction, or Office macro security to make the helper start.

## Register the startup task

Create one task for the employee who is installing PPTK.

| Setting | Value |
|---|---|
| Name | `PPTKKeys Startup` |
| Trigger | At logon of that user only |
| Principal | The same user |
| Logon type | Interactive token. Run only when the user is logged on |
| Run level | Limited. Highest privileges stay off |
| Multiple instances | Ignore a new start while the task is already running |
| Time limit | None (`PT0S`) |
| Stop when idle ends | Off |
| Action | `"<dotnet.exe>" exec "<dll>"` |
| Working directory | The directory that contains `PPTKKeys.dll` |

Example action on this machine:

```text
C:\Users\NoufAlDakheel\AppData\Local\Microsoft\dotnet\dotnet.exe exec "C:\Users\NoufAlDakheel\AppData\Local\PPTK\PPTKKeys\PPTKKeys.dll"
```

Working directory:

```text
C:\Users\NoufAlDakheel\AppData\Local\PPTK\PPTKKeys
```

`dotnet.exe` is a console program, so Windows opens a console window for it at logon (in Windows Terminal when that is the default terminal). PPTKKeys detaches from that console as soon as it starts, when it is the only process attached, so the window closes after a short flash. Started from a terminal, the console is shared and stays. A helper with no flash at all needs a Windows-subsystem `PPTKKeys.exe`; the ASR rule "Block executable files from running unless they meet a prevalence, age, or trusted list criterion" blocks an unsigned one, so that requires a trusted code signature or an IT decision.

The application also holds the local mutex `Local\PPTKKeys.Windows`. A second copy exits with "PPTKKeys is already running." The task policy and the mutex both prevent two hooks.

Do not add a Run-key or Startup-folder entry as a substitute. If Task Scheduler rejects the registration, report the policy or permission error and stop.

## Update the helper

1. `Stop-ScheduledTask -TaskName "PPTKKeys Startup"`. Ending the task stops the process it started.
2. Wait until no `PPTKKeys.exe` process remains and no command line contains `PPTKKeys.dll`.
3. Replace the three files in `%LOCALAPPDATA%\PPTK\PPTKKeys`.
4. `Start-ScheduledTask -TaskName "PPTKKeys Startup"`.
5. Confirm exactly one process, and confirm its loaded `PPTKKeys.dll` is the file just installed.

Do not start the new copy before the old process has exited. The mutex would make the new copy quit, and the old hook would keep running.

## Uninstall

1. Stop the task, then wait until the helper process is gone.
2. `Unregister-ScheduledTask -TaskName "PPTKKeys Startup" -Confirm:$false`
3. Delete `%LOCALAPPDATA%\PPTK\PPTKKeys`.
4. Leave the user's .NET runtime in place unless this installer was the component that installed it.
5. Do not remove `PPTK.ppam` as part of the helper uninstall unless the same installer owns that add-in.

## Detect a missing runtime later

On each install and update:

1. Parse `framework.name` and `framework.version` from `PPTKKeys.runtimeconfig.json`.
2. Run `"<dotnet.exe>" --list-runtimes`.
3. Treat a same-major, equal-or-higher patch as present. `10.0.0` is present when `10.0.14` is listed.
4. If `dotnet.exe` itself is missing, or the listed runtimes do not satisfy the file, stop and tell the employee which framework to install.
5. A failed `dotnet exec` whose output says the framework was not found is the same condition. Do not register the startup task in that case.
