VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   Caption         =   "Sjoerd File Upload"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7410
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   7410
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox F 
      Height          =   1455
      Left            =   3840
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSWinsockLib.Winsock w 
      Index           =   0
      Left            =   840
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wSock 
      Left            =   360
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   80
   End
   Begin VB.Image imgExitA 
      Height          =   615
      Left            =   240
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label lblExit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      Height          =   195
      Left            =   720
      TabIndex        =   3
      Top             =   3960
      Width           =   255
   End
   Begin VB.Image imgExit 
      Height          =   360
      Left            =   240
      Picture         =   "frmMain.frx":0442
      Top             =   3840
      Width           =   345
   End
   Begin VB.Image imgSettingsA 
      Height          =   615
      Left            =   240
      Top             =   840
      Width           =   1335
   End
   Begin VB.Image imgStopserver 
      Height          =   615
      Left            =   240
      Top             =   240
      Width           =   1575
   End
   Begin VB.Image imgStartserver 
      Height          =   615
      Left            =   240
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Image imgServer 
      Height          =   390
      Left            =   240
      Picture         =   "frmMain.frx":0B44
      Top             =   360
      Width           =   360
   End
   Begin VB.Label lblServer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Stop server"
      Height          =   195
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   810
   End
   Begin VB.Label lblSettings 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      Height          =   195
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   570
   End
   Begin VB.Image imgSettings 
      Height          =   420
      Left            =   240
      Picture         =   "frmMain.frx":100A
      Top             =   960
      Width           =   390
   End
   Begin VB.Image imgBack 
      Height          =   4455
      Left            =   0
      Picture         =   "frmMain.frx":1521
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public UseLogin As Boolean
Public UseAudio As Boolean
Public UseDocument As Boolean
Public UsePicture As Boolean
Public UseSoftware As Boolean
Public UseRest As Boolean
Public UseForum As Boolean
Private ServerState As String
Private Temp()

Private Sub Form_Load()
On Error Resume Next
ReDim Temp(500)

ServerFolder = mdlINI.mfncGetFromIni("General", "Serverfolder", App.Path & "\Settings.ini")
strURL = mdlINI.mfncGetFromIni("General", "URL", App.Path & "\Settings.ini")
UseForum = mdlINI.mfncGetFromIni("General", "Forum", App.Path & "\Settings.ini")

UseAudio = mdlINI.mfncGetFromIni("Folders", "Audio", App.Path & "\Settings.ini")
UsePicture = mdlINI.mfncGetFromIni("Folders", "Picture", App.Path & "\Settings.ini")
UseDocument = mdlINI.mfncGetFromIni("Folders", "Document", App.Path & "\Settings.ini")
UseSoftware = mdlINI.mfncGetFromIni("Folders", "Software", App.Path & "\Settings.ini")
UseRest = mdlINI.mfncGetFromIni("Folders", "Rest", App.Path & "\Settings.ini")

UseLogin = mdlINI.mfncGetFromIni("Security", "Password protection", App.Path & "\Settings.ini")

DownloadFolder = ServerFolder & "\Program files"
str404 = "<HTML>" & vbCrLf & "<BODY>" & vbCrLf & _
"<p>Sorry, the requested page was not found<br>" & vbCrLf & _
"Perhaps you 've misspelled the URL or the URL is no longer availble.</p>" & vbCrLf & _
"<p>Please check the home page:<br>" & vbCrLf & _
"<a href=""http://" & strURL & """>" & vbCrLf & _
"http://" & strURL & "</a></p>" & vbCrLf & _
"<hr>" & vbCrLf & _
"<p>Sjoerd Huininga<br>" & vbCrLf & _
"Copyright 2004<br>" & vbCrLf & _
"&nbsp;</p>" & vbCrLf & _
"</BODY>" & vbCrLf & _
"</HTML>"
str403 = "<HTML>" & vbCrLf & "<BODY>" & vbCrLf & _
"<p>Sorry, you're not allowed to view the requested page<br>" & vbCrLf & _
"Perhaps you need some special permission</p>" & vbCrLf & _
"<p>Please check the home page:<br>" & vbCrLf & _
"<a href=""http://" & strURL & """>" & vbCrLf & _
"http://" & strURL & "</a></p>" & vbCrLf & _
"<hr>" & vbCrLf & _
"<p>Sjoerd Huininga<br>" & vbCrLf & _
"Copyright 2004<br>" & vbCrLf & _
"&nbsp;</p>" & vbCrLf & _
"</BODY>" & vbCrLf & _
"</HTML>"
StartServer
End Sub

Private Sub imgBack_Click()
MsgBox "Current site: " & strURL & "" & vbCrLf & "Current server state: " & ServerState, , "Sjoerd Server"
End Sub

Private Sub imgExitA_Click()
On Error Resume Next
Dim I As Integer
For I = 0 To w.Count - 1
w(I).Close
DoEvents
Next
End
End Sub

Private Sub imgSettingsA_Click()
frmSettings.Show
End Sub

Private Sub imgStartserver_Click()
StartServer
imgStopserver.Visible = True
imgStartserver.Visible = False
lblServer.Caption = "Stop server"
End Sub

Private Sub imgStopserver_Click()
StopServer
imgStartserver.Visible = True
imgStopserver.Visible = False
lblServer.Caption = "Start server"
End Sub

Private Sub StartServer()
On Error Resume Next
wSock.Listen
ServerState = "Online"
End Sub
Private Sub StopServer()
On Error Resume Next
wSock.Close
ServerState = "Offline"
End Sub

Private Sub w_DataArrival(Index As Integer, ByVal bytesTotal As Long)
On Error Resume Next

Dim aGa As String
Dim aKa As String
Dim aFa As String
Dim strFile As String
Dim strBoundary As String
Dim strRaw As String
Dim strData As String
Dim IP As String
Dim BAN As String
Dim intF As Integer
Dim UserName As String
Dim Co

w(Index).GetData strData

Debug.Print strData & "Remote IP: " & w(Index).RemoteHostIP

IP = mdlINI.mfncGetFromIni("VISITORS", w(Index).RemoteHostIP, ServerFolder & "\IPData.ini")
BAN = mdlINI.mfncGetFromIni("BAN", w(Index).RemoteHostIP, ServerFolder & "\IP.ini")

If BAN = "Yes" Then
    w(Index).SendData mdlMain.Banned(w(Index).RemoteHostIP)
    Exit Sub
End If

Co = Split(strData, " HTTP/1.1")
aGa = Co(0)
aKa = Mid$(aGa, 5)
aKa = Replace(aKa, "/", "\")
aKa = Replace(aKa, "%20", " ")

If InStr(1, strData, "forum=") Then 'the program now see's that someone submitted txt to the forum
    w(Index).SendData mdlMain.PostMessage(strData, w(Index).RemoteHostIP)
    Exit Sub
End If

If InStr(1, strData, "Vak=") Then 'never mind the syntax i'm checking; it's a souvenir of an old project
    w(Index).SendData mdlMain.PostHuiswerk(strData)
    Exit Sub
End If

If InStr(1, strData, "Register=") Then 'i gues this is easy to understand :p
    w(Index).SendData mdlMain.Register(strData, w(Index).RemoteHostIP)
    Exit Sub
End If

If InStr(1, strData, "Username=") Then 'and so is this
    w(Index).SendData mdlMain.Login(strData, w(Index).RemoteHostIP)
    Exit Sub
End If

If aKa = "\regist1.gif" Then 'if you don't do this, u can't use these Pictures before you're logged in
    w(Index).SendData LoadPage(aKa, w(Index).RemoteHost, w(Index).RemoteHost)
    Exit Sub
End If

If aKa = "\homepa8.gif" Then 'same as above
    w(Index).SendData LoadPage(aKa, w(Index).RemoteHost, w(Index).RemoteHost)
    Exit Sub
End If

If aKa = "\ashton13.jpg" Then 'you have to change this to the path of the background from your register/login page
    w(Index).SendData LoadPage(aKa, w(Index).RemoteHost, w(Index).RemoteHost)
    Exit Sub
End If

If UseLogin = True Then 'if you want to use login the following lines are important
If Not IP = "Yes" Then 'check if you already have an account registrered on this ip adress
    w(Index).SendData mdlMain.FirstTime 'send the welcome message...
    Exit Sub
End If
End If

If UseLogin = True Then 'same as above
If mdlINI.mfncGetFromIni(w(Index).RemoteHostIP, "Login", ServerFolder & "\login.ini") = "False" Then
    w(Index).SendData mdlMain.LoadPage("\verification.htm", w(Index).RemoteHostIP, w(Index).RemoteHostIP)
    Exit Sub
End If
End If

If InStr(1, strData, "/download?custom") Then 'custom homepage; now uses c:\custom.jpg as background
    w(Index).SendData mdlMain.HomePageCustom(strData)
    Exit Sub
End If

If InStr(1, strData, "/logout") Then 'log a user out; handy if you have a dynamic ip; so nobody else can acces your account trough his/her ip
    w(Index).SendData mdlMain.Logout(w(Index).RemoteHostIP)
    Exit Sub
End If

If aKa = "\" Then aKa = "\download" 'if no page then homepage

If UseAudio = True Then 'if you want to use the audio option the following lines are important
If aKa = "\audio" Then
    w(Index).SendData mdlMain.LoadAudio(w(Index).RemoteHost)
    Exit Sub
End If
End If

If UseDocument = True Then 'same as above
If aKa = "\document" Then
    w(Index).SendData mdlMain.LoadDocument(w(Index).RemoteHost)
    Exit Sub
End If
End If

If UsePicture = True Then 'same as above
If aKa = "\Picture" Then
    w(Index).SendData mdlMain.loadPicture(w(Index).RemoteHost)
    Exit Sub
End If
End If

If UseSoftware = True Then  'same as above
If aKa = "\software" Then
    w(Index).SendData mdlMain.LoadSoftware(w(Index).RemoteHost)
    Exit Sub
End If
End If

If UseRest = True Then 'same as above
If aKa = "\Rest" Then
    w(Index).SendData mdlMain.loadRest(w(Index).RemoteHost)
    Exit Sub
End If
End If

If UseForum = True Then 'same as above
If aKa = "\forum" Then
    w(Index).SendData mdlMain.LoadPage("\forum.htm", w(Index).RemoteHost, w(Index).LocalIP)
    Exit Sub
End If
End If

If aKa = "\download" Then 'sends home page
    w(Index).SendData mdlMain.HomePage(w(Index).RemoteHost)
    Exit Sub
End If

w(Index).SendData mdlMain.LoadPage(aKa, w(Index).LocalIP, w(Index).RemoteHostIP) 'load any file requested
End Sub

Private Sub w_SendComplete(Index As Integer)
On Error Resume Next
w(Index).Close
End Sub

Private Sub wSock_ConnectionRequest(ByVal requestID As Long)
On Error Resume Next
Dim aAa As Integer
aAa = w.Count
Load w(aAa)
w(aAa).Accept requestID
End Sub

