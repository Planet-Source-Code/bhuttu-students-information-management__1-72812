VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form frmRestore 
   BorderStyle     =   0  'None
   Caption         =   "Restore Database"
   ClientHeight    =   5895
   ClientLeft      =   1590
   ClientTop       =   4320
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   2130
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin prjStudentIMS.ucGradContainer fraBackGround 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   10398
      IconSize        =   0
      Caption         =   "Restore Database"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   4080
         TabIndex        =   6
         Top             =   1455
         Width           =   2715
      End
      Begin VB.DirListBox Dir1 
         Height          =   1890
         Left            =   1380
         TabIndex        =   5
         Top             =   1440
         Width           =   2565
      End
      Begin VB.DriveListBox Drive1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1380
         TabIndex        =   4
         Top             =   1080
         Width           =   5445
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Restore"
         Default         =   -1  'True
         Height          =   360
         Left            =   4515
         TabIndex        =   2
         Top             =   5370
         WhatsThisHelpID =   3056
         Width           =   1140
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         Height          =   360
         Left            =   5715
         TabIndex        =   1
         Top             =   5370
         WhatsThisHelpID =   3055
         Width           =   1140
      End
      Begin prjStudentIMS.ucHorizontal3DLine ucHorizontal3DLine1 
         Height          =   30
         Left            =   375
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   5220
         Width           =   6465
         _ExtentX        =   11404
         _ExtentY        =   53
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   1335
         TabIndex        =   7
         Top             =   4770
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   240
         Picture         =   "frmRestore.frx":0000
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label lblMsg 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1365
         TabIndex        =   11
         Top             =   4440
         Width           =   5445
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select the backed up database file and click the Restore button below"
         Height          =   195
         Index           =   2
         Left            =   1365
         TabIndex        =   10
         Top             =   720
         Width           =   5010
      End
      Begin VB.Label lblFileName 
         BackStyle       =   0  'Transparent
         Height          =   510
         Left            =   1365
         TabIndex        =   9
         Top             =   3820
         Width           =   5430
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Selected databse backup file and the location path"
         Height          =   195
         Index           =   0
         Left            =   1365
         TabIndex        =   8
         Top             =   3480
         Width           =   3630
      End
   End
End
Attribute VB_Name = "frmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sFileName               As String
'

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdOk_Click()
  
  
End Sub

Private Sub Dir1_Change()
    lblFileName.Caption = ""
On Error GoTo A1:
    File1.Path = Dir1.Path
    Exit Sub
A1:
    MsgBox "Folder can not be accessed ...", vbInformation, "Drive not accessible ..."
    Drive1.Drive = "c:"
End Sub

Private Sub Drive1_Change()
    lblFileName.Caption = ""
On Error GoTo A1:
    Dir1.Path = Drive1.Drive
    Exit Sub
A1:
    MsgBox "Drive can not be accessed ...", vbInformation, "Drive not accessible ..."
    Drive1.Drive = "c:"

End Sub

Private Sub File1_Click()
    sFileName = File1.FileName
    lblFileName.Caption = File1.Path & "\" & File1.FileName
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        If ProgressBar1.Value = 0 Or ProgressBar1.Value Then
            Call Unload(Me)
        End If
    End If
End Sub

Private Sub Form_Load()
    Call CenterForm(Me)
    Icon = frmMain.Icon
    
    Call initFrame(fraBackGround)
    
    lblFileName.Caption = ""
    lblMsg.Visible = False
    ProgressBar1.Visible = False
    
    Drive1.Drive = "c:"
    
End Sub


