Option Explicit

Dim objShell, objFSO, strScriptPath, strRepoPath
Dim strOutput, intResult

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Ermittle den Ordner, in dem das Skript liegt
strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
strRepoPath = strScriptPath

' Wechsle ins Repository-Verzeichnis
objShell.CurrentDirectory = strRepoPath

' Pull ausführen
Dim objExec
Set objExec = objShell.Exec("git pull")

' Warte auf Fertigstellung
Do While objExec.Status = 0
    WScript.Sleep 100
Loop

strOutput = objExec.StdOut.ReadAll()

If objExec.ExitCode = 0 Then
    If InStr(strOutput, "Already up to date") > 0 Then
        MsgBox "Repository ist bereits aktuell.", vbInformation, "Git Pull"
    Else
        MsgBox "Erfolgreich aktualisiert!" & vbCrLf & vbCrLf & strOutput, vbInformation, "Git Pull"
    End If
Else
    MsgBox "Fehler beim Pull:" & vbCrLf & vbCrLf & objExec.StdErr.ReadAll(), vbCritical, "Git Pull"
End If

Set objExec = Nothing
Set objShell = Nothing
Set objFSO = Nothing