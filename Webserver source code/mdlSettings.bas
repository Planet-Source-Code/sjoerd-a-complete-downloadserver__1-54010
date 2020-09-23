Attribute VB_Name = "mdlSettings"
Option Explicit

Public Sub SetAll()
On Error Resume Next

ServerFolder = frmSettings.txtServerfolder

mdlINI.mfncWriteIni "General", "Serverfolder", frmSettings.txtServerfolder, App.Path & "\Settings.ini"
mdlINI.mfncWriteIni "General", "URL", frmSettings.txtURL, App.Path & "\Settings.ini"
mdlINI.mfncWriteIni "General", "Forum", frmMain.UseForum, App.Path & "\Settings.ini"

mdlINI.mfncWriteIni "Folders", "Audio", frmMain.UseAudio, App.Path & "\Settings.ini"
mdlINI.mfncWriteIni "Folders", "Picture", frmMain.UsePicture, App.Path & "\Settings.ini"
mdlINI.mfncWriteIni "Folders", "Document", frmMain.UseDocument, App.Path & "\Settings.ini"
mdlINI.mfncWriteIni "Folders", "Software", frmMain.UseSoftware, App.Path & "\Settings.ini"
mdlINI.mfncWriteIni "Folders", "Rest", frmMain.UseRest, App.Path & "\Settings.ini"

MkDir ServerFolder & "\Audio"
MkDir ServerFolder & "\Picture"
MkDir ServerFolder & "\Document"
MkDir ServerFolder & "\Software"
MkDir ServerFolder & "\Rest"

mdlINI.mfncWriteIni "Security", "Password protection", frmMain.UseLogin, App.Path & "\Settings.ini"
End Sub
