Option Explicit

Dim objShell, objFSO, strScriptPath, strRepoPath
Dim strStatus, strMessage, intResult

Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Ermittle den Ordner, in dem das Skript liegt
strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
strRepoPath = strScriptPath

' Wechsle ins Repository-Verzeichnis
objShell.CurrentDirectory = strRepoPath

' Zeige Status
strStatus = objShell.Exec("git status --short").StdOut.ReadAll()

If Len(Trim(strStatus)) = 0 Then
    MsgBox "Keine Änderungen gefunden.", vbInformation, "Git Commit"
    WScript.Quit
End If

MsgBox "Geänderte Dateien:" & vbCrLf & vbCrLf & strStatus, vbInformation, "Git Status"

' Frage nach Commit-Beschreibung
strMessage = InputBox("Commit-Beschreibung eingeben:", "Git Commit")

If Trim(strMessage) = "" Then
    MsgBox "Keine Beschreibung eingegeben. Abbruch.", vbExclamation, "Git Commit"
    WScript.Quit
End If

' Stage alle Änderungen
objShell.Run "cmd /c git add .", 0, True

' Commit
intResult = objShell.Run("cmd /c git commit -m """ & strMessage & """", 0, True)

If intResult <> 0 Then
    MsgBox "Fehler beim Commit!", vbCritical, "Git Commit"
    WScript.Quit
End If

' Push
intResult = objShell.Run("cmd /c git push", 0, True)

If intResult = 0 Then
    MsgBox "Erfolgreich committed und gepusht!", vbInformation, "Git Commit"
Else
    MsgBox "Fehler beim Push! Bitte manuell prüfen.", vbCritical, "Git Push"
End If

Set objShell = Nothing
Set objFSO = Nothing