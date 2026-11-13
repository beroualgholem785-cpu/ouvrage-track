' Ouvrage Track - lanceur (ouvre l'application dans sa propre fenetre, sans barre de navigateur)
Option Explicit
Dim sh, fso, here, appDir, idx, data, exe, cands, c, url, i
Set sh = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
appDir = here & "\app"
idx = appDir & "\index.html"
If Not fso.FileExists(idx) Then
  MsgBox "Fichier introuvable : " & idx, 16, "Ouvrage Track"
  WScript.Quit 1
End If
' Les donnees (marches, situations...) sont gardees sur CET ordinateur
data = sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\OuvrageTrack\profil"
If Not fso.FolderExists(sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\OuvrageTrack") Then fso.CreateFolder sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\OuvrageTrack"
If Not fso.FolderExists(data) Then fso.CreateFolder data

' Dossier de sauvegarde automatique : Documents\Sauvegardes Ouvrage Track
Dim bk, pref, f, js
bk = sh.SpecialFolders("MyDocuments") & "\Sauvegardes Ouvrage Track"
If Not fso.FolderExists(bk) Then fso.CreateFolder bk
' Le navigateur depose les sauvegardes dans ce dossier, sans demander ou les enregistrer
If Not fso.FolderExists(data & "\Default") Then fso.CreateFolder data & "\Default"
pref = data & "\Default\Preferences"
If Not fso.FileExists(pref) Then
  js = Replace(bk, "\", "\\")
  Set f = fso.CreateTextFile(pref, True)
  f.Write "{""download"":{""default_directory"":""" & js & """,""prompt_for_download"":false,""directory_upgrade"":true},""savefile"":{""default_directory"":""" & js & """}}"
  f.Close
End If
cands = Array( _
  sh.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Microsoft\Edge\Application\msedge.exe", _
  sh.ExpandEnvironmentStrings("%ProgramFiles%") & "\Microsoft\Edge\Application\msedge.exe", _
  sh.ExpandEnvironmentStrings("%ProgramFiles%") & "\Google\Chrome\Application\chrome.exe", _
  sh.ExpandEnvironmentStrings("%ProgramFiles(x86)%") & "\Google\Chrome\Application\chrome.exe", _
  sh.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Google\Chrome\Application\chrome.exe")
exe = ""
For Each c In cands
  If fso.FileExists(c) Then exe = c : Exit For
Next
If exe = "" Then
  MsgBox "Microsoft Edge ou Google Chrome est necessaire." & vbCrLf & "Installez Google Chrome puis relancez Ouvrage Track.", 48, "Ouvrage Track"
  WScript.Quit 1
End If
url = "file:///" & Replace(Replace(idx, "\", "/"), " ", "%20")
sh.Run """" & exe & """ --app=""" & url & """ --user-data-dir=""" & data & """ --allow-file-access-from-files --no-first-run --no-default-browser-check --window-size=1400,900", 1, False
