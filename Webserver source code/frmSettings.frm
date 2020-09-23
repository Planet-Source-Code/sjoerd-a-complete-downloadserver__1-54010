VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Settings"
   ClientHeight    =   4965
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   6285
   Icon            =   "frmSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   3
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample4 
         Caption         =   "Sample 4"
         Height          =   1785
         Left            =   2100
         TabIndex        =   11
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3780
      Index           =   2
      Left            =   -20000
      ScaleHeight     =   3780
      ScaleWidth      =   5685
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSample3 
         Caption         =   "Sample 3"
         Height          =   1785
         Left            =   1545
         TabIndex        =   10
         Top             =   675
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   4560
      Width           =   1095
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   3585
      Index           =   1
      Left            =   -20000
      ScaleHeight     =   3585
      ScaleWidth      =   5685
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraConnections 
         Caption         =   "Connections"
         Height          =   1905
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   5295
         Begin VB.TextBox txtAmount 
            Height          =   285
            Left            =   3240
            TabIndex        =   23
            Top             =   240
            Width           =   1935
         End
         Begin VB.CheckBox chkLimit 
            Caption         =   "Check to set no limit at all"
            Height          =   255
            Left            =   240
            TabIndex        =   22
            Top             =   480
            Width           =   2415
         End
         Begin VB.Label lblMax 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Max. amount of connections at same time:"
            Height          =   195
            Left            =   240
            TabIndex        =   21
            Top             =   240
            Width           =   2985
         End
      End
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   4140
      Index           =   0
      Left            =   240
      ScaleHeight     =   4140
      ScaleWidth      =   5685
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   480
      Width           =   5685
      Begin VB.Frame fraSecurity 
         Caption         =   "Security"
         Height          =   1185
         Left            =   240
         TabIndex        =   19
         Top             =   2880
         Width           =   5055
         Begin VB.CheckBox chkSecurity 
            Caption         =   "Users have to register and login before they can use the site"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   4815
         End
      End
      Begin VB.Frame fraFolders 
         Caption         =   "Folders"
         Height          =   2505
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   5055
         Begin VB.TextBox txtURL 
            Height          =   285
            Left            =   1080
            TabIndex        =   25
            Top             =   600
            Width           =   2295
         End
         Begin VB.CheckBox chkForum 
            Caption         =   "Use forum"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   2160
            Width           =   2055
         End
         Begin VB.CheckBox chkRest 
            Caption         =   "Use rest folder"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1920
            Width           =   2055
         End
         Begin VB.CheckBox chkSoftware 
            Caption         =   "Use software folder"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1680
            Width           =   2055
         End
         Begin VB.CheckBox chkDocument 
            Caption         =   "Use document folder"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1440
            Width           =   1815
         End
         Begin VB.CheckBox chkPicture 
            Caption         =   "Use picture folder"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1200
            Width           =   1695
         End
         Begin VB.CheckBox chkAudio 
            Caption         =   "Use audio folder"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox txtServerfolder 
            Height          =   285
            Left            =   1080
            TabIndex        =   13
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label lblUrl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "URL:"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   375
         End
         Begin VB.Label lblServerfolder 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Serverfolder: "
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   945
         End
      End
   End
   Begin MSComctlLib.TabStrip tbsOptions 
      Height          =   4605
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8123
      HotTracking     =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "Group1"
            Object.ToolTipText     =   "General settings"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Some easy codes to temporary save the clicked settings.
Private Sub chkAudio_Click()
    If chkAudio.Value = 1 Then
    frmMain.UseAudio = True
    Else
    frmMain.UseAudio = False
    End If
End Sub

Private Sub chkPicture_Click()
    If chkPicture.Value = 1 Then
    frmMain.UsePicture = True
    Else
    frmMain.UsePicture = False
    End If
End Sub

Private Sub chkDocument_Click()
    If chkDocument.Value = 1 Then
    frmMain.UseDocument = True
    Else
    frmMain.UseDocument = False
    End If
End Sub

Private Sub chkSecurity_Click()
    If chkSecurity.Value = 1 Then
    frmMain.UseLogin = True
    Else
    frmMain.UseLogin = False
    End If
End Sub

Private Sub chkSoftware_Click()
    If chkSoftware.Value = 1 Then
    frmMain.UseSoftware = True
    Else
    frmMain.UseSoftware = False
    End If
End Sub

Private Sub chkRest_Click()
    If chkRest.Value = 1 Then
    frmMain.UseRest = True
    Else
    frmMain.UseRest = False
    End If
End Sub

Private Sub chkForum_Click()
    If chkForum.Value = 1 Then
    frmMain.UseForum = True
    Else
    frmMain.UseForum = False
    End If
End Sub

Private Sub cmdApply_Click()
    mdlSettings.SetAll
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    mdlSettings.SetAll
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim I As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        I = tbsOptions.SelectedItem.Index
        If I = tbsOptions.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(1)
        Else
            'increment the tab
            Set tbsOptions.SelectedItem = tbsOptions.Tabs(I + 1)
        End If
    End If
End Sub

Private Sub Form_Load()
On Error Resume Next
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
    txtServerfolder.Text = ServerFolder
    txtURL.Text = strURL
    
    If frmMain.UseAudio = True Then chkAudio.Value = 1
    If frmMain.UsePicture = True Then chkPicture.Value = 1
    If frmMain.UseDocument = True Then chkDocument.Value = 1
    If frmMain.UseSoftware = True Then chkSoftware.Value = 1
    If frmMain.UseRest = True Then chkRest.Value = 1
    If frmMain.UseForum = True Then chkForum.Value = 1
    If frmMain.UseLogin = True Then chkSecurity.Value = 1
End Sub

Private Sub fraSample1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub tbsOptions_Click()
    
    Dim I As Integer
    'show and enable the selected tab's controls
    'and hide and disable all others
    For I = 0 To tbsOptions.Tabs.Count - 1
        If I = tbsOptions.SelectedItem.Index - 1 Then
            picOptions(I).Left = 210
            picOptions(I).Enabled = True
        Else
            picOptions(I).Left = -20000
            picOptions(I).Enabled = False
        End If
    Next
    
End Sub
