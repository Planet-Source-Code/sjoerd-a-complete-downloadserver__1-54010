Attribute VB_Name = "mdlMain"
Option Explicit

Public ServerFolder As String
Public DownloadFolder As String
Public str404 As String
Public str403 As String
Public strURL As String
Global NRs As Integer

Public Function Banned(IP$) As String
On Error GoTo errHandler
Dim Warning As String
Warning = "<p><font color=""#FF0000""" & ">" & vbCrLf & _
"<marquee>Warning: You are banned, you broke the rules" & vbCrLf & _
"</marquee></font></p>"
Banned = "<TITLE>You are banned from " & strURL & "</TITLE>" & vbCrLf & _
"The server has banned the following IP adress: " & IP$ & "<Br>" & _
"Because this IP has broke the rules on my site" & vbCrLf & Warning

Exit Function

errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
Banned = str404
End If
End Function

Public Function FirstTime() As String
On Error GoTo errHandler
Dim strData As String
Dim intF As Integer
intF = FreeFile
Open DownloadFolder & "\registration.htm" For Binary As #1
strData = Space$(LOF(intF))
Get #1, , strData
Close #1
FirstTime = strData

Exit Function

errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
FirstTime = str404
End If
End Function

Public Function HomePage(IP$) As String
On Error GoTo errHandler
Dim strAudio As String
Dim strAudioS As String
Dim strPicture As String
Dim strPictureS As String
Dim strDocument As String
Dim strDocumentS As String
Dim strSoftware As String
Dim strSoftwareS As String
Dim strRest As String
Dim strRestS As String
Dim strForum As String
Dim strData As String

frmMain.F.Path = ServerFolder & "\Audio"
frmMain.F.Refresh
strAudio = frmMain.F.ListCount
If frmMain.UseAudio = True Then strAudioS = "<p><a href=/audio><img border=0 src=../homepa1.gif width=32 height=36> All audio (" & strAudio & ")</p>"

frmMain.F.Path = ServerFolder & "\Picture"
frmMain.F.Refresh
strPicture = frmMain.F.ListCount
If frmMain.UsePicture = True Then strPictureS = "<p><a href=/Picture><img border=0 src=../homepa2.gif width=29 height=31> All Pictures" & vbCrLf & _
"(" & strPicture & ")</p>"

frmMain.F.Path = ServerFolder & "\Document"
frmMain.F.Refresh
strDocument = frmMain.F.ListCount
If frmMain.UseDocument = True Then strDocumentS = "<p><a href=/document><img border=0 src=../homepa3.gif width=30 height=32> All documents (" & strDocument & ")</p>"

frmMain.F.Path = ServerFolder & "\Software"
frmMain.F.Refresh
strSoftware = frmMain.F.ListCount
If frmMain.UseSoftware = True Then strSoftwareS = "<p><a href=/software><img border=0 src=../homepa4.jpg width=31 height=33> All software (" & strSoftware & ")</p>"

frmMain.F.Path = ServerFolder & "\Rest"
frmMain.F.Refresh
strRest = frmMain.F.ListCount
If frmMain.UseRest = True Then strRestS = "<p><a href=/Rest><img border=0 src=../homepa5.jpg width=30 height=34> Rest (" & strRest & ")</p>"

If frmMain.UseForum = True Then strForum = "<p><a href=/forum><img border=0 src=../homepa6.gif width=26 height=28> Forum</p>"

strData = "<html>" & vbCrLf & _
"<head>" & vbCrLf & _
"<meta http-equiv=Content-Language content=nl>" & vbCrLf & _
"<meta name=GENERATOR content=Microsoft FrontPage 5.0>" & vbCrLf & _
"<meta name=ProgId content=FrontPage.Editor.Document>" & vbCrLf & _
"<meta http-equiv=Content-Type content=text/html; charset=windows-1252>" & vbCrLf & _
"<title>Sjoerd File Upload</title>" & vbCrLf & _
"</head>" & vbCrLf & _
"<body background=../ashton13.jpg link=#000000 vlink=#000000 alink=#000000>" & vbCrLf & _
"<p>Sjoerd File Upload</p>" & vbCrLf & _
"<p>&nbsp;</p>" & vbCrLf & strAudioS & vbCrLf & strPictureS & vbCrLf & strDocumentS & vbCrLf & _
strSoftwareS & vbCrLf & strRestS & vbCrLf & _
"</br></br>" & vbCrLf & _
"</br></br>" & vbCrLf & _
"</br></br>" & vbCrLf & _
"</br></br>" & vbCrLf & _
"</br></br>" & vbCrLf & _
strForum
strData = strData & vbCrLf & _
"<p><a href=/logout>" & vbCrLf & _
"<img border=0 src=../regist1.gif width=24 height=26> Logout</a></p>" & vbCrLf & _
"</body>" & vbCrLf & _
"</html>" & vbCrLf
HomePage = strData
Exit Function
errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
HomePage = str404
End If
End Function
Public Function HomePageCustom(IP$) As String
On Error GoTo errHandler
Dim strAudio As String
Dim strAudioS As String
Dim strPicture As String
Dim strPictureS As String
Dim strDocument As String
Dim strDocumentS As String
Dim strSoftware As String
Dim strSoftwareS As String
Dim strRest As String
Dim strRestS As String
Dim strForum As String
Dim strData As String

frmMain.F.Path = ServerFolder & "\Audio"
frmMain.F.Refresh
strAudio = frmMain.F.ListCount
If frmMain.UseAudio = True Then strAudioS = "<p><a href=/audio><img border=0 src=../homepa1.gif width=32 height=36> All audio (" & strAudio & ")</p>"

frmMain.F.Path = ServerFolder & "\Picture"
frmMain.F.Refresh
strPicture = frmMain.F.ListCount
If frmMain.UsePicture = True Then strPictureS = "<p><a href=/Picture><img border=0 src=../homepa2.gif width=29 height=31> All Pictures" & vbCrLf & _
"(" & strPicture & ")</p>"

frmMain.F.Path = ServerFolder & "\Document"
frmMain.F.Refresh
strDocument = frmMain.F.ListCount
If frmMain.UseDocument = True Then strDocumentS = "<p><a href=/document><img border=0 src=../homepa3.gif width=30 height=32> All documents (" & strDocument & ")</p>"

frmMain.F.Path = ServerFolder & "\Software"
frmMain.F.Refresh
strSoftware = frmMain.F.ListCount
If frmMain.UseSoftware = True Then strSoftwareS = "<p><a href=/software><img border=0 src=../homepa4.jpg width=31 height=33> All software (" & strSoftware & ")</p>"

frmMain.F.Path = ServerFolder & "\Rest"
frmMain.F.Refresh
strRest = frmMain.F.ListCount
If frmMain.UseRest = True Then strRestS = "<p><a href=/Rest><img border=0 src=../homepa5.jpg width=30 height=34> Rest (" & strRest & ")</p>"

If frmMain.UseForum = True Then strForum = "<p><a href=/forum><img border=0 src=../homepa6.gif width=26 height=28> Forum</p>"

strData = "<html>" & vbCrLf & _
"<head>" & vbCrLf & _
"<meta http-equiv=Content-Language content=nl>" & vbCrLf & _
"<meta name=GENERATOR content=Microsoft FrontPage 5.0>" & vbCrLf & _
"<meta name=ProgId content=FrontPage.Editor.Document>" & vbCrLf & _
"<meta http-equiv=Content-Type content=text/html; charset=windows-1252>" & vbCrLf & _
"<title>Sjoerd File Upload</title>" & vbCrLf & _
"</head>" & vbCrLf & _
"<body background=C:\custom.jpg link=#000000 vlink=#000000 alink=#000000>" & vbCrLf & _
"<p>Sjoerd File Upload</p>" & vbCrLf & _
"<p>&nbsp;</p>" & vbCrLf & strAudioS & vbCrLf & strPictureS & vbCrLf & strDocumentS & vbCrLf & _
strSoftwareS & vbCrLf & strRestS & vbCrLf & _
"</br></br>" & vbCrLf & _
"</br></br>" & vbCrLf & _
"</br></br>" & vbCrLf & _
"</br></br>" & vbCrLf & _
"</br></br>" & vbCrLf & _
strForum
strData = strData & vbCrLf & _
"<p><a href=/logout>" & vbCrLf & _
"<img border=0 src=../regist1.gif width=24 height=26> Logout</a></p>" & vbCrLf & _
"</body>" & vbCrLf & _
"</html>" & vbCrLf
HomePageCustom = strData
Exit Function
errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
HomePageCustom = str404
End If
End Function
Public Function Logout(IP$) As String
On Error GoTo errHandler
Dim UserName As String
UserName = mdlINI.mfncGetFromIni(IP$, "Username", ServerFolder & "\login.ini")
mdlINI.mfncWriteIni IP$, "Login", "False", ServerFolder & "\login.ini"
Logout = Confirmation(UserName)
Exit Function
errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
Logout = str404
End If
End Function

Public Function Confirmation(UserName As String) As String
On Error GoTo errHandler
Confirmation = "<title>Logout succeeded</title>" & vbCrLf & _
"</head>" & vbCrLf & _
"<body>" & vbCrLf & _
"<hr>" & vbCrLf & _
"<p><span lang=en-us>Log gemaakt: <font color=#008000><b>Ok</b><br>" & vbCrLf & _
"</font>Username en password: <b><font color=#008000>Ok<br>" & vbCrLf & _
"<br>" & vbCrLf & _
"</font></b>U wordt over een paar seconden naar de hoofdpagina geleid.</span></p>" & vbCrLf & _
"<hr>"
Exit Function
errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
Confirmation = str404
End If
End Function

Public Function LoadAudio(IP$) As String
On Error GoTo errHandler
Dim strData As String
Dim strFiles As String
Dim strAudio As String
Dim I As Integer

frmMain.F.Path = ServerFolder & "\Audio"
frmMain.F.Refresh
strAudio = frmMain.F.ListCount

For I = 0 To strAudio - 1
strFiles = strFiles & _
"<p><a title=""Category: Documenten"" href=""/Audio/" & frmMain.F.List(I) & """>" & Mid$(frmMain.F.List(I), 1, Len(frmMain.F.List(I)) - 4) & " </a></p>"
Next

strData = "<html>" & vbCrLf & _
"<head>" & vbCrLf & _
"<meta http-equiv=Content-Language content=nl>" & vbCrLf & _
"<meta name=GENERATOR content=Microsoft FrontPage 5.0>" & vbCrLf & _
"<meta name=ProgId content=FrontPage.Editor.Document>" & vbCrLf & _
"<meta http-equiv=Content-Type content=text/html; charset=windows-1252>" & vbCrLf & _
"<title>Sjoerd File Upload</title>" & vbCrLf & _
"</head>" & vbCrLf & _
"<body background=ashton13.jpg link=#000000 vlink=#000000 alink=#000000>" & vbCrLf & _
"<p>Sjoerd File Upload</p>" & vbCrLf & _
"<p>&nbsp;</p>" & vbCrLf & _
"<p><img border=0 src=homepa1.gif width=32 height=36>Audio<br>" & vbCrLf & _
"<a href=/download>" & vbCrLf & _
"<img border=0 src=homepa8.gif width=29 height=28>Home</a></p>" & vbCrLf & _
strFiles

LoadAudio = strData

Exit Function
errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
LoadAudio = str404
End If
End Function

Public Function loadPicture(IP$) As String
Dim strData As String
Dim strFiles As String
Dim strPicture As String
Dim Size As Long
Dim intF As Integer
Dim I As Integer

frmMain.F.Path = ServerFolder & "\Picture"
frmMain.F.Refresh
strPicture = frmMain.F.ListCount

For I = 0 To strPicture - 1
DoEvents
strFiles = strFiles & _
"<p><a title=""Category: Pictures"" href=""/Picture/" & frmMain.F.List(I) & """>" & Mid$(frmMain.F.List(I), 1, Len(frmMain.F.List(I)) - 4) & " </a></p>"
Next

strData = "<html>" & vbCrLf & _
"<head>" & vbCrLf & _
"<meta http-equiv=Content-Language content=nl>" & vbCrLf & _
"<meta name=GENERATOR content=Microsoft FrontPage 5.0>" & vbCrLf & _
"<meta name=ProgId content=FrontPage.Editor.Document>" & vbCrLf & _
"<meta http-equiv=Content-Type content=text/html; charset=windows-1252>" & vbCrLf & _
"<title>Sjoerd File Upload</title>" & vbCrLf & _
"</head>" & vbCrLf & _
"<body background=ashton13.jpg link=#000000 vlink=#000000 alink=#000000>" & vbCrLf & _
"<p>Sjoerd File Upload</p>" & vbCrLf & _
"<p>&nbsp;</p>" & vbCrLf & _
"<p><img border=0 src=homepa2.gif>Picture<br>" & vbCrLf & _
"<a href=/download>" & vbCrLf & _
"<img border=0 src=homepa8.gif width=29 height=28>Home</a></p>" & vbCrLf & _
strFiles

loadPicture = strData

Exit Function
errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
loadPicture = str404
End If
End Function

Public Function LoadDocument(IP$) As String
Dim strData As String
Dim strFiles As String
Dim strDocument As String
Dim I As Integer

frmMain.F.Path = ServerFolder & "\Document"
frmMain.F.Refresh
strDocument = frmMain.F.ListCount

For I = 0 To strDocument - 1
strFiles = strFiles & _
"<p><a title=""Category: Documenten"" href=""/Document/" & frmMain.F.List(I) & """>" & Mid$(frmMain.F.List(I), 1, Len(frmMain.F.List(I)) - 4) & " </a></p>"
Next

strData = "<html>" & vbCrLf & _
"<head>" & vbCrLf & _
"<meta http-equiv=Content-Language content=nl>" & vbCrLf & _
"<meta name=GENERATOR content=Microsoft FrontPage 5.0>" & vbCrLf & _
"<meta name=ProgId content=FrontPage.Editor.Document>" & vbCrLf & _
"<meta http-equiv=Content-Type content=text/html; charset=windows-1252>" & vbCrLf & _
"<title>Sjoerd File Upload</title>" & vbCrLf & _
"</head>" & vbCrLf & _
"<body background=ashton13.jpg link=#000000 vlink=#000000 alink=#000000>" & vbCrLf & _
"<p>Sjoerd File Upload</p>" & vbCrLf & _
"<p>&nbsp;</p>" & vbCrLf & _
"<p><img border=0 src=homepa3.gif>Document<br>" & vbCrLf & _
"<a href=/download>" & vbCrLf & _
"<img border=0 src=homepa8.gif width=29 height=28>Home</a></p>" & vbCrLf & _
strFiles

LoadDocument = strData

Exit Function
errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
LoadDocument = str404
End If
End Function

Public Function LoadSoftware(IP$) As String
Dim strData As String
Dim strFiles As String
Dim strDocument As String
Dim I As Integer

frmMain.F.Path = ServerFolder & "\Software"
frmMain.F.Refresh
strDocument = frmMain.F.ListCount

For I = 0 To strDocument - 1
strFiles = strFiles & _
"<p><a title=""Category: Software"" href=""/Software/" & frmMain.F.List(I) & """>" & Mid$(frmMain.F.List(I), 1, Len(frmMain.F.List(I)) - 4) & " </a></p>"
Next

strData = "<html>" & vbCrLf & _
"<head>" & vbCrLf & _
"<meta http-equiv=Content-Language content=nl>" & vbCrLf & _
"<meta name=GENERATOR content=Microsoft FrontPage 5.0>" & vbCrLf & _
"<meta name=ProgId content=FrontPage.Editor.Document>" & vbCrLf & _
"<meta http-equiv=Content-Type content=text/html; charset=windows-1252>" & vbCrLf & _
"<title>Sjoerd File Upload</title>" & vbCrLf & _
"</head>" & vbCrLf & _
"<body background=ashton13.jpg link=#000000 vlink=#000000 alink=#000000>" & vbCrLf & _
"<p>Sjoerd File Upload</p>" & vbCrLf & _
"<p>&nbsp;</p>" & vbCrLf & _
"<p><img border=0 src=homepa4.jpg>Software<br>" & vbCrLf & _
"<a href=/download>" & vbCrLf & _
"<img border=0 src=homepa8.gif width=29 height=28>Home</a></p>" & vbCrLf & _
strFiles

LoadSoftware = strData

Exit Function
errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
LoadSoftware = str404
End If
End Function

Public Function loadRest(IP$) As String
Dim strData As String
Dim strFiles As String
Dim strDocument As String
Dim I As Integer

frmMain.F.Path = ServerFolder & "\Rest"
frmMain.F.Refresh
strDocument = frmMain.F.ListCount

For I = 0 To strDocument - 1
strFiles = strFiles & _
"<p><a title=""Category: Rest"" href=""/Rest/" & frmMain.F.List(I) & """>" & Mid$(frmMain.F.List(I), 1, Len(frmMain.F.List(I)) - 4) & " </a></p>"
Next

strData = "<html>" & vbCrLf & _
"<head>" & vbCrLf & _
"<meta http-equiv=Content-Language content=nl>" & vbCrLf & _
"<meta name=GENERATOR content=Microsoft FrontPage 5.0>" & vbCrLf & _
"<meta name=ProgId content=FrontPage.Editor.Document>" & vbCrLf & _
"<meta http-equiv=Content-Type content=text/html; charset=windows-1252>" & vbCrLf & _
"<title>Sjoerd File Upload</title>" & vbCrLf & _
"</head>" & vbCrLf & _
"<body background=ashton13.jpg link=#000000 vlink=#000000 alink=#000000>" & vbCrLf & _
"<p>Sjoerd File Upload</p>" & vbCrLf & _
"<p>&nbsp;</p>" & vbCrLf & _
"<p><img border=0 src=homepa5.jpg>Rest<br>" & vbCrLf & _
"<a href=/download>" & vbCrLf & _
"<img border=0 src=homepa8.gif width=29 height=28>Home</a></p>" & vbCrLf & _
strFiles

loadRest = strData

Exit Function
errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
loadRest = str404
End If
End Function

Public Function PostMessage(aKa, IP$)

On Error Resume Next

Dim strData As String
Dim strNaam As String
Dim strNaamd
Dim strNaams
Dim strBericht As String
Dim strBerichtd
Dim strBerichts
Dim strMail As String
Dim strMaild
Dim strMails
Dim strAvater As String
Dim strAvaterd
Dim strAvaters
Dim UserName As String

aKa = Replace(aKa, "%0D%0A", "<BR>")
aKa = Replace(aKa, "+", " ")
aKa = Replace(aKa, "%2C", ",")
aKa = Replace(aKa, "%21", "!")
aKa = Replace(aKa, "%23", "#")
aKa = Replace(aKa, "%24", "$")
aKa = Replace(aKa, "%25", "%")
aKa = Replace(aKa, "%5E", "^")
aKa = Replace(aKa, "%28", "(")
aKa = Replace(aKa, "%29", ")")
aKa = Replace(aKa, "%2B", "+")
aKa = Replace(aKa, "%5C", "\")
aKa = Replace(aKa, "%7C", "|")
aKa = Replace(aKa, "%2F", "/")
aKa = Replace(aKa, "%3F", "?")
aKa = Replace(aKa, "%3C", "<")
aKa = Replace(aKa, "%3E", ">")
aKa = Replace(aKa, "%3A", ":")
aKa = Replace(aKa, "%27", "'")
aKa = Replace(aKa, "%22", """")
aKa = Replace(aKa, "%3B", ";")
aKa = Replace(aKa, "%5B", "[")
aKa = Replace(aKa, "%5D", "]")
aKa = Replace(aKa, "%7D", "}")
aKa = Replace(aKa, "%7B", "{")
aKa = Replace(aKa, "%60", "`")
aKa = Replace(aKa, "%7E", "~")

strBericht = Mid$(aKa, InStr(1, aKa, "Bericht"))
strNaam = Mid$(aKa, InStr(1, aKa, "forum"))
strMail = Mid$(aKa, InStr(1, aKa, "email"))
strAvater = Mid$(aKa, InStr(1, aKa, "avater"))

strNaamd = Split(strNaam, "=")
strNaams = Split(strNaamd(1), "&")

strBerichtd = Split(strBericht, "=")
strBerichts = Split(strBerichtd(1), "&")

strMaild = Split(strMail, "=")
strMails = Split(strMaild(1), "&")

strAvaterd = Split(strAvater, "=")
strAvaters = Split(strAvaterd(1), "&")

strBericht = Left$(strBerichts(0), Len(strBerichts(0)))
strNaam = Left$(strNaams(0), Len(strNaams(0)))
strMail = Left$(strMails(0), Len(strMails(0)))
strAvater = Left$(strAvaters(0), Len(strAvaters(0)))

strBericht = Replace(strBericht, "%26", "&")
strBericht = Replace(strBericht, "%3D", "=")
strBericht = Replace(strBericht, ">:)", "<img border=0 src=.\Pictures\emoticon.bmp>")
strBericht = Replace(strBericht, ":)", "<img border=0 src=.\Pictures\blij.bmp>")
strBericht = Replace(strBericht, ";)", "<img border=0 src=.\Pictures\knipoog.bmp>")
strBericht = Replace(strBericht, ":o", "<img border=0 src=.\Pictures\verrast.bmp>")
strBericht = Replace(strBericht, ":p", "<img border=0 src=.\Pictures\tong.bmp>")
strBericht = Replace(strBericht, "(h)", "<img border=0 src=.\Pictures\zonnebril.bmp>")
strBericht = Replace(strBericht, ":@", "<img border=0 src=.\Pictures\verhit.bmp>")
strBericht = Replace(strBericht, ":s", "<img border=0 src=.\Pictures\vaag.bmp>")
strBericht = Replace(strBericht, ":$", "<img border=0 src=.\Pictures\schaam.bmp")
strBericht = Replace(strBericht, ":(", "<img border=0 src=.\Pictures\jammer.bmp>")
strBericht = Replace(strBericht, ":'(", "<img border=0 src=.\Pictures\huilen.bmp>")
strBericht = Replace(strBericht, ":|", "<img border=0 src=.\Pictures\verbaasd.bmp>")
strBericht = Replace(strBericht, "(a)", "<img border=0 src=.\Pictures\engel.bmp>")
strBericht = Replace(strBericht, ":|", "<img border=0 src=.\Pictures\verbaasd.bmp>")
strBericht = Replace(strBericht, "8o|", "<img border=0 src=.\Pictures\tanden.bmp>")
strBericht = Replace(strBericht, "8-|", "<img border=0 src=.\Pictures\bril.bmp>")
strBericht = Replace(strBericht, "+o(", "<img border=0 src=.\Pictures\ziek.bmp>")
strBericht = Replace(strBericht, "<:o)|", "<img border=0 src=.\Pictures\feest.bmp>")
strBericht = Replace(strBericht, "|-)", "<img border=0 src=.\Pictures\slaap.bmp>")
strBericht = Replace(strBericht, "*-)", "<img border=0 src=.\Pictures\denken.bmp>")
strBericht = Replace(strBericht, ":-#", "<img border=0 src=.\Pictures\niemand.bmp>")
strBericht = Replace(strBericht, ":-*", "<img border=0 src=.\Pictures\geheim.bmp>")
strBericht = Replace(strBericht, "^o)", "<img border=0 src=.\Pictures\sarcasme.bmp>")
strBericht = Replace(strBericht, "8-)", "<img border=0 src=.\Pictures\rol.bmp>")
strBericht = Replace(strBericht, "(l)", "<img border=0 src=.\Pictures\hart.bmp>")
strBericht = Replace(strBericht, "(u)", "<img border=0 src=.\Pictures\gebroken hart.bmp>")
strBericht = Replace(strBericht, "(m)", "<img border=0 src=.\Pictures\msn.bmp>")
strBericht = Replace(strBericht, "(@)", "<img border=0 src=.\Pictures\kat.bmp>")
strBericht = Replace(strBericht, "(&)", "<img border=0 src=.\Pictures\hond.bmp>")
strBericht = Replace(strBericht, "(sn)", "<img border=0 src=.\Pictures\slak.bmp>")
strBericht = Replace(strBericht, "(bah)", "<img border=0 src=.\Pictures\slak.bmp>")
strBericht = Replace(strBericht, ":d", "<img border=0 src=.\Pictures\lachend.bmp>")
strBericht = Replace(strBericht, "(sn)", "<img border=0 src=.\Pictures\slak.bmp>")
strBericht = Replace(strBericht, "(l)", "<img border=0 src=.\Pictures\hart.bmp>")

If strBericht = Empty Then Exit Function
UserName = mdlINI.mfncGetFromIni(IP$, "Username", ServerFolder & "\login.ini")
strData = "<table border=1 cellpadding=0 cellspacing=0 style=""border-collapse: collapse"" bordercolor=#111111 width=100% id=AutoNumber1 background=""" & UserName & """>" & vbCrLf & _
"  <tr>" & vbCrLf & _
    "<td width=34% valign=top>Naam: " & strNaam & "<BR>" & vbCrLf & _
    "E-mail: " & strMail & " op: " & Now & "<BR>" & vbCrLf & _
     "<img border=0 src=""" & strAvater & """ width=337 height=233></td>" & vbCrLf & _
    "<td width=66% valign=top><span lang=nl>Bericht:</span><p>" & strBericht & vbCrLf & _
    "</td>" & vbCrLf & _
  "</tr>" & vbCrLf & _
"</table>" & vbCrLf & _
"<hr>"

Open DownloadFolder & "\forum.htm" For Append As #1
Print #1, strData
Close #1

PostMessage = Bevestiging1
End Function

Public Function LoadPage(Filename As String, IP$, IP1$) As String
On Error GoTo errHandler
Dim FileName1 As String
Dim strHeader As String
Dim Cont As String
Dim F As Integer
Dim textda As String

F = FreeFile
textda = ""
FileName1 = Filename
Filename = DownloadFolder & Filename

If Not InStr(1, FileName1, "\Audio\") = 0 Then
Filename = ServerFolder & FileName1
End If
If Not InStr(1, FileName1, "\Picture\") = 0 Then
Filename = ServerFolder & FileName1
End If
If Not InStr(1, FileName1, "\Document\") = 0 Then
Filename = ServerFolder & FileName1
End If
If Not InStr(1, FileName1, "\Software\") = 0 Then
Filename = ServerFolder & FileName1
End If
If Not InStr(1, FileName1, "\Rest\") = 0 Then
Filename = ServerFolder & FileName1
End If

If FileExists(Filename) Then
    If Len(Filename) Then
        Open Filename For Binary As #F
            textda = Space$(LOF(F))
            Get #F, , textda
            DoEvents
        Close #F
    End If
        Select Case LCase(Mid(Filename, InStrRev(Filename, ".") + 1))
        Case "html" 'HTML file
            Cont = "text/html"
        Case "htm"  'HTML file
            Cont = "text/html"
        Case "txt"  'TEXT (notepad) file
            Cont = "text/text"
        Case "js"   'Javascript library
            Cont = "text/html" 'YES, that is right
        Case "pdf"  'ADOBE ACROBAT PDF file
            Cont = "application/pdf"
        Case "sit"  'STUFFIT archive
            Cont = "application/x-stuffit"
        Case "avi"  'AUDIO VISUAL video
            Cont = "video/avi"
        Case "asf"
            Cont = "video/mpeg"
        Case "css"  'CASSCADING STYLE SHEET formating info
            Cont = "text/css"
        Case "swf"  'SHOCKWAVE FLASH animation
            Cont = "application/futuresplash"
        Case "jpg"  'JOINT PHOTOGROPHERS EXPERT GROUP image
            Cont = "image/jpeg"
        Case "xls"  'MICROSOFT EXCEL spreadsheet
            Cont = "application/vnd.ms-excel"
        Case "doc"  'MICROSOFT WORD formated text
            Cont = "application/vnd.ms-word"
        Case "midi" 'MUSICAL INSTRUMENT DIGITAL INTERFACE music
            Cont = "audio/midi"
        Case "mp3"  'MOTION Picture EXPERT GROUP LAYER 3 music
            Cont = "audio/mpeg"
        Case "wma"  'MOTION Picture EXPERT GROUP LAYER 3 music"
            Cont = "audio/mpeg"
        Case "rm"   'REAL MEDIA video
            Cont = "application/vnd.rn-realmedia"
        Case "rtf"  'MICROSOFT RICHTEXT formatted text
            Cont = "application/msword"
        Case "exe"
            Cont = "application/octet"
        Case "wav"  'WAVE sound
            Cont = "audio/wav"
        Case "zip"  'ZIP archive
            Cont = "application/x-zip"
        Case "png"  'PORTABLE NETWORK GRAPGHICS image
            Cont = "image/png"
        Case "gif"  'COMPUSERVE GRAPHICS INTERGHANGE FORMAT Image
            Cont = "image/gif"
        End Select
strHeader = "HTTP/1.1 200 OK" & vbCrLf & _
"Server: Sjoerd Webserver" & vbCrLf & _
"Host: " & IP$ & vbCrLf & _
"Accept-Ranges-: bytes" & vbCrLf & _
"Content-Length: " & Len(textda) & vbCrLf & _
"Connection: Close" & vbCrLf & _
"Content-Type: " & Cont

LoadPage = strHeader & vbCrLf & vbCrLf & textda

Else
Err.Raise 53
End If
Exit Function

errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
LoadPage = str404
Exit Function
End If
LoadPage = "<TITLE>Sjoerd File Upload</TITLE>" & vbCrLf & _
"Sorry, there was a server based error"
End Function

Public Function FileExists(ByVal sFileName As String) As Integer
'Checks if the given file exists.

Dim intLen As Integer
On Error Resume Next

    intLen = Len(Dir$(sFileName))
    
    If Err Or intLen = 0 Then
        FileExists = False
        Else
            FileExists = True
    End If
End Function

Public Function Login(aKa As String, IP$) As String

On Error GoTo errHandler


Dim strData As String
Dim strUsername As String
Dim strUsernamed
Dim strUsernames
Dim strPassword As String
Dim strPasswordd
Dim strPasswords
Dim strMail As String
Dim strMaild
Dim strMails

aKa = Replace(aKa, "%0D%0A", "<BR>")
aKa = Replace(aKa, "+", " ")
aKa = Replace(aKa, "%2C", ",")
aKa = Replace(aKa, "%21", "!")
aKa = Replace(aKa, "%23", "#")
aKa = Replace(aKa, "%24", "$")
aKa = Replace(aKa, "%25", "%")
aKa = Replace(aKa, "%5E", "^")
aKa = Replace(aKa, "%28", "(")
aKa = Replace(aKa, "%29", ")")
aKa = Replace(aKa, "%2B", "+")
aKa = Replace(aKa, "%5C", "\")
aKa = Replace(aKa, "%7C", "|")
aKa = Replace(aKa, "%2F", "/")
aKa = Replace(aKa, "%3F", "?")
aKa = Replace(aKa, "%3C", "<")
aKa = Replace(aKa, "%3E", ">")
aKa = Replace(aKa, "%3A", ":")
aKa = Replace(aKa, "%27", "'")
aKa = Replace(aKa, "%22", """")
aKa = Replace(aKa, "%3B", ";")
aKa = Replace(aKa, "%5B", "[")
aKa = Replace(aKa, "%5D", "]")
aKa = Replace(aKa, "%7D", "}")
aKa = Replace(aKa, "%7B", "{")
aKa = Replace(aKa, "%60", "`")
aKa = Replace(aKa, "%7E", "~")

strUsername = Mid$(aKa, InStr(1, aKa, "Username"))
strPassword = Mid$(aKa, InStr(1, aKa, "Password"))
strMail = Mid$(aKa, InStr(1, aKa, "Password1"))

strUsernamed = Split(strUsername, "=")
strUsernames = Split(strUsernamed(1), "&")

strPasswordd = Split(strPassword, "=")
strPasswords = Split(strPasswordd(1), "&")

strMaild = Split(strMail, "=")
strMails = Split(strMaild(1), "&")

strPassword = Left$(strPasswords(0), Len(strPasswords(0)))
strUsername = Left$(strUsernames(0), Len(strUsernames(0)))
strMail = Left$(strMails(0), Len(strMails(0)))

If strPassword = "" Or strUsername = "" Then Login = LoadPage("\verification.htm", IP$, IP$): Exit Function

If Not strPassword = strMail Then Login = LoadPage("\verification.htm", IP$, IP$): Exit Function

If strPassword = mdlINI.mfncGetFromIni(strUsername, "Password", ServerFolder & "\register.ini") Then
Login = Router(strUsername, IP$)
Else
Login = Forbidden
End If
Exit Function

errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
Login = str404
Exit Function
End If
Login = "<TITLE>Sjoerd File Upload</TITLE>" & vbCrLf & _
"Sorry, there was a server based error"
End Function

Public Function Register(aKa As String, IP$) As String

On Error GoTo errHandler

Dim strData As String
Dim strUsername As String
Dim strUsernamed
Dim strUsernames
Dim strPassword As String
Dim strPasswordd
Dim strPasswords
Dim strMail As String
Dim strMaild
Dim strMails

aKa = Replace(aKa, "%0D%0A", "<BR>")
aKa = Replace(aKa, "+", " ")
aKa = Replace(aKa, "%2C", ",")
aKa = Replace(aKa, "%21", "!")
aKa = Replace(aKa, "%23", "#")
aKa = Replace(aKa, "%24", "$")
aKa = Replace(aKa, "%25", "%")
aKa = Replace(aKa, "%5E", "^")
aKa = Replace(aKa, "%28", "(")
aKa = Replace(aKa, "%29", ")")
aKa = Replace(aKa, "%2B", "+")
aKa = Replace(aKa, "%5C", "\")
aKa = Replace(aKa, "%7C", "|")
aKa = Replace(aKa, "%2F", "/")
aKa = Replace(aKa, "%3F", "?")
aKa = Replace(aKa, "%3C", "<")
aKa = Replace(aKa, "%3E", ">")
aKa = Replace(aKa, "%3A", ":")
aKa = Replace(aKa, "%27", "'")
aKa = Replace(aKa, "%22", """")
aKa = Replace(aKa, "%3B", ";")
aKa = Replace(aKa, "%5B", "[")
aKa = Replace(aKa, "%5D", "]")
aKa = Replace(aKa, "%7D", "}")
aKa = Replace(aKa, "%7B", "{")
aKa = Replace(aKa, "%60", "`")
aKa = Replace(aKa, "%7E", "~")
aKa = Replace(aKa, "%3D", "=")

strUsername = Mid$(aKa, InStr(1, aKa, "Register"))
strPassword = Mid$(aKa, InStr(1, aKa, "Password"))
strMail = Mid$(aKa, InStr(1, aKa, "email"))

strUsernamed = Split(strUsername, "=")
strUsernames = Split(strUsernamed(1), "&")

strPasswordd = Split(strPassword, "=")
strPasswords = Split(strPasswordd(1), "&")

strMaild = Split(strMail, "=")
strMails = Split(strMaild(1), "&")

strPassword = Left$(strPasswords(0), Len(strPasswords(0)))
strUsername = Left$(strUsernames(0), Len(strUsernames(0)))
strMail = Left$(strMails(0), Len(strMails(0)))

If strPassword = "" Or strUsername = "" Or strMail = "" Then Register = LoadPage("\registration.htm", IP$, IP$): Exit Function

If strUsername = "download" Then Register = str404: Exit Function
If strUsername = "admin" Then Register = str404: Exit Function

mdlINI.WritePrivateProfileString strUsername, "UserName", strUsername, ServerFolder & "\register.ini"
mdlINI.WritePrivateProfileString strUsername, "Password", strPassword, ServerFolder & "\register.ini"
mdlINI.WritePrivateProfileString strUsername, "Email", strMail, ServerFolder & "\register.ini"

On Error Resume Next

mdlINI.mfncWriteIni strUsername, "Password", strPassword, ServerFolder & "\register.ini"
mdlINI.mfncWriteIni strUsername, "Email", strMail, ServerFolder & "\register.ini"
mdlINI.mfncWriteIni "VISITORS", IP$, "Yes", ServerFolder & "\IPData.ini"

Register = Bevestiging2
Exit Function

errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
Register = str404
Exit Function
End If
Register = "<TITLE>Sjoerd File Upload</TITLE>" & vbCrLf & _
"Sorry, there was a server based error"
End Function

Public Function PostHuiswerk(aKa As String) As String

On Error GoTo errHandler

Dim strData As String
Dim strNaam As String
Dim strNaamd
Dim strNaams
Dim strHuiswerk As String
Dim strHuiswerkd
Dim strHuiswerks
Dim strMail As String
Dim strMaild
Dim strMails
Dim strVak As String
Dim strVakd
Dim strVaks

aKa = Replace(aKa, "%0D%0A", "<BR>")
aKa = Replace(aKa, "+", " ")
aKa = Replace(aKa, "%2C", ",")
aKa = Replace(aKa, "%21", "!")
aKa = Replace(aKa, "%23", "#")
aKa = Replace(aKa, "%24", "$")
aKa = Replace(aKa, "%25", "%")
aKa = Replace(aKa, "%5E", "^")
aKa = Replace(aKa, "%28", "(")
aKa = Replace(aKa, "%29", ")")
aKa = Replace(aKa, "%2B", "+")
aKa = Replace(aKa, "%5C", "\")
aKa = Replace(aKa, "%7C", "|")
aKa = Replace(aKa, "%2F", "/")
aKa = Replace(aKa, "%3F", "?")
aKa = Replace(aKa, "%3C", "<")
aKa = Replace(aKa, "%3E", ">")
aKa = Replace(aKa, "%3A", ":")
aKa = Replace(aKa, "%27", "'")
aKa = Replace(aKa, "%22", """")
aKa = Replace(aKa, "%3B", ";")
aKa = Replace(aKa, "%5B", "[")
aKa = Replace(aKa, "%5D", "]")
aKa = Replace(aKa, "%7D", "}")
aKa = Replace(aKa, "%7B", "{")
aKa = Replace(aKa, "%60", "`")
aKa = Replace(aKa, "%7E", "~")

strVak = Mid$(aKa, InStr(1, aKa, "Vak"))
strHuiswerk = Mid$(aKa, InStr(1, aKa, "Huiswerk"))
strNaam = Mid$(aKa, InStr(1, aKa, "naam"))
strMail = Mid$(aKa, InStr(1, aKa, "email"))

strNaamd = Split(strNaam, "=")
strNaams = Split(strNaamd(1), "&")

strVakd = Split(strVak, "=")
strVaks = Split(strVakd(1), "&")

strHuiswerkd = Split(strHuiswerk, "=")
strHuiswerks = Split(strHuiswerkd(1), "&")

strMaild = Split(strMail, "=")
strMails = Split(strMaild(1), "&")

strVak = Left$(strVaks(0), Len(strVaks(0)))
strHuiswerk = Left$(strHuiswerks(0), Len(strHuiswerks(0)))
strNaam = Left$(strNaams(0), Len(strNaams(0)))
strMail = Left$(strMails(0), Len(strMails(0)))

strHuiswerk = Replace(strHuiswerk, "%26", "&")
strHuiswerk = Replace(strHuiswerk, "%3D", "=")

Open ServerFolder & "\Document\" & strNaam & "-" & strVak & ".htm" For Binary As #1
Put #1, , strHuiswerk & "<BR>"
Put #1, , "Gepost door: " & strMail & " op: " & Now
Close #1

PostHuiswerk = Bevestiging
Exit Function

errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
PostHuiswerk = str404
Exit Function
End If
PostHuiswerk = "<TITLE>Sjoerd File Upload</TITLE>" & vbCrLf & _
"Sorry, there was a server based error"
End Function

Public Function Forbidden() As String
On Error GoTo errHandler
Dim strData As String
Dim intF As Integer
intF = FreeFile
strData = Empty
Open ServerFolder & "\forbidden.htm" For Binary As #1
strData = Space$(LOF(intF))
Get #1, , strData
Close #1
Forbidden = strData
strData = Empty
Exit Function

errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
Forbidden = str404
Exit Function
End If
Forbidden = "<TITLE>Sjoerd File Upload</TITLE>" & vbCrLf & _
"Sorry, there was a server based error"
End Function

Public Function Router(UserName As String, IP$) As String
On Error GoTo errHandler
Dim strData As String

Router = "<meta http-equiv=refresh content=5;url=\>" & vbCrLf & _
"<title>Login succeeded</title>" & vbCrLf & _
"</head>" & vbCrLf & _
"<body>" & vbCrLf & _
"<hr>" & vbCrLf & _
"<p><span lang=en-us>Log gemaakt: <font color=#008000><b>Ok</b><br>" & vbCrLf & _
"</font>Username en password: <b><font color=#008000>Ok<br>" & vbCrLf & _
"<br>" & vbCrLf & _
"</font></b>You're now being transfered to the home page.</span></p>" & vbCrLf & _
"<hr>"
mdlINI.mfncWriteIni UserName, Now, "Time", ServerFolder & "\login.ini"
mdlINI.mfncWriteIni IP$, "Username", UserName, ServerFolder & "\login.ini"
mdlINI.mfncWriteIni IP$, "Login", "True", ServerFolder & "\login.ini"
strData = Empty
Exit Function

errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
Router = str404
Exit Function
End If
Router = "<TITLE>Sjoerd File Upload</TITLE>" & vbCrLf & _
"Sorry, there was a server based error"
End Function

Public Function Bevestiging2() As String
On Error GoTo errHandler
Dim strData As String, intF As Integer
intF = FreeFile
strData = Empty
Open DownloadFolder & "\bedankt2.htm" For Binary As #1
strData = Space$(LOF(intF))
Get #1, , strData
Close #1
Bevestiging2 = strData
strData = Empty
Exit Function

errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
Bevestiging2 = str404
Exit Function
End If
Bevestiging2 = "<TITLE>Sjoerd File Upload</TITLE>" & vbCrLf & _
"Sorry, there was a server based error"
End Function

Public Function Bevestiging1() As String
On Error GoTo errHandler
Dim strData As String, intF As Integer
intF = FreeFile
strData = Empty
Open DownloadFolder & "\bedankt1.htm" For Binary As #1
strData = Space$(LOF(intF))
Get #1, , strData
Close #1
Bevestiging1 = strData
strData = Empty
Exit Function

errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
Bevestiging1 = str404
Exit Function
End If
Bevestiging1 = "<TITLE>Sjoerd File Upload</TITLE>" & vbCrLf & _
"Sorry, there was a server based error"
End Function

Public Function Bevestiging() As String
On Error GoTo errHandler
Dim strData As String, intF As Integer
intF = FreeFile
strData = Empty
Open DownloadFolder & "\bedankt.htm" For Binary As #1
strData = Space$(LOF(intF))
Get #1, , strData
Close #1
Bevestiging = strData
strData = Empty
Exit Function

errHandler:
If Err.Number = 53 Or Err.Number = 76 Then
Bevestiging = str404
Exit Function
End If
Bevestiging = "<TITLE>Sjoerd File Upload</TITLE>" & vbCrLf & _
"Sorry, there was a server based error"
End Function

