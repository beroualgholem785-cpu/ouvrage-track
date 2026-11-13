' Ouvrage Track - installation sur cet ordinateur
Option Explicit
Dim sh, fso, here, dest, desk, lnk, sm, src
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
dest = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\OuvrageTrack"
If Not fso.FolderExists(dest) Then fso.CreateFolder dest
If fso.FolderExists(dest & "\app") Then fso.DeleteFolder dest & "\app", True
fso.CopyFolder here & "\app", dest & "\app", True
fso.CopyFile here & "\Lancer.vbs", dest & "\Lancer.vbs", True
fso.CopyFile here & "\ouvrage-track.ico", dest & "\ouvrage-track.ico", True
Sub MakeLink(path)
  Set lnk = sh.CreateShortcut(path)
  lnk.TargetPath = sh.ExpandEnvironmentStrings("%SystemRoot%") & "\System32\wscript.exe"
  lnk.Arguments = """" & dest & "\Lancer.vbs"""
  lnk.WorkingDirectory = dest
  lnk.IconLocation = dest & "\ouvrage-track.ico"
  lnk.Description = "Ouvrage Track"
  lnk.Save
End Sub
Dim bk, l2
bk = sh.SpecialFolders("MyDocuments") & "\Sauvegardes Ouvrage Track"
If Not fso.FolderExists(bk) Then fso.CreateFolder bk
Set l2 = sh.CreateShortcut(sh.SpecialFolders("Desktop") & "\Sauvegardes Ouvrage Track.lnk")
l2.TargetPath = bk
l2.Description = "Sauvegardes automatiques Ouvrage Track"
l2.Save
MakeLink sh.SpecialFolders("Desktop") & "\Ouvrage Track.lnk"
MakeLink sh.SpecialFolders("Programs") & "\Ouvrage Track.lnk"
MsgBox "Ouvrage Track est installe." & vbCrLf & vbCrLf & "Cliquez sur l'icone Ouvrage Track sur le Bureau pour ouvrir l'application." & vbCrLf & "Vos sauvegardes automatiques seront dans Documents\Sauvegardes Ouvrage Track." & vbCrLf & "Vous pouvez retirer la cle USB.", 64, "Ouvrage Track"
sh.Run "wscript.exe """ & dest & "\Lancer.vbs""", 1, False
