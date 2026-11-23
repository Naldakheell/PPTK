Attribute VB_Name = "ModulePPTKHelperStart"
Option Explicit

' Starts the PPTKKeys keyboard-shortcut helper when PowerPoint loads PPTK and the helper is not running
' (for example after its console window was closed at sign-in, which ends the helper).
'
' Called from PPTK_RibbonOnLoad (ModulePPTKAlign). Windows only; silent: any problem (no task, no access,
' WMI unavailable) leaves everything as it is and shows nothing.
'   1. The task "\PPTKKeys Startup" is running: the helper runs. Nothing to do (fast path).
'   2. An install, update or uninstall holds %LOCALAPPDATA%\PPTK\state.lock: do not start the helper while
'      install.ps1 replaces its files.
'   3. No dotnet.exe running PPTKKeys.dll (a helper started by an update handoff runs outside the task): the
'      task is run through Task Scheduler (Schedule.Service). The helper is then started by Task Scheduler,
'      not by PowerPoint, and its own mutex keeps a second helper from running.
' No Win32 Declare (blocked by the ASR rule on Win32 imports from Office macros).

Private Const HelperTaskName As String = "PPTKKeys Startup"
Private Const TaskStateRunning As Long = 4

Public Sub PPTK_EnsureHelper()
#If Mac Then
    Exit Sub
#Else
    Dim scheduler As Object
    Dim helperTask As Object

    On Error GoTo Quiet
    Set scheduler = CreateObject("Schedule.Service")
    scheduler.Connect
    Set helperTask = scheduler.GetFolder("\").GetTask(HelperTaskName)
    If helperTask.State = TaskStateRunning Then Exit Sub
    If PPTK_DeployBusy() Then Exit Sub
    If PPTK_HelperRunning() Then Exit Sub
    helperTask.Run vbNullString
Quiet:
#End If
End Sub

#If Not Mac Then
' True when install.ps1 or uninstall.ps1 holds state.lock (opened without sharing for the whole run).
Private Function PPTK_DeployBusy() As Boolean
    Dim lockPath As String
    Dim fileNumber As Integer

    lockPath = Environ$("LOCALAPPDATA") & "\PPTK\state.lock"
    If Len(Dir$(lockPath, vbHidden Or vbSystem)) = 0 Then Exit Function
    fileNumber = FreeFile
    On Error GoTo Busy
    Open lockPath For Binary Access Read Shared As #fileNumber
    Close #fileNumber
    Exit Function
Busy:
    PPTK_DeployBusy = True
End Function

' True when a dotnet.exe process runs PPTKKeys.dll (the helper, started by the task or by an update handoff).
Private Function PPTK_HelperRunning() As Boolean
    Dim wmi As Object
    Dim process As Object

    Set wmi = GetObject("winmgmts:")
    For Each process In wmi.ExecQuery("SELECT CommandLine FROM Win32_Process WHERE Name = 'dotnet.exe'")
        If InStr(1, process.CommandLine & "", "PPTKKeys.dll", vbTextCompare) > 0 Then
            PPTK_HelperRunning = True
            Exit Function
        End If
    Next process
End Function
#End If
