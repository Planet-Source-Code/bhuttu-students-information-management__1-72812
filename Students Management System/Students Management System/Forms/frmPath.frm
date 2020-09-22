VERSION 5.00
Begin VB.Form frmPath 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4815
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjStudentIMS.ucGradContainer fraBackGround 
      Height          =   4815
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   8493
      Caption         =   "Select Databse Backup Location"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3660
         TabIndex        =   7
         Top             =   4320
         WhatsThisHelpID =   3055
         Width           =   1140
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2400
         TabIndex        =   6
         Top             =   4320
         WhatsThisHelpID =   3056
         Width           =   1140
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   4440
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   360
         TabIndex        =   1
         Top             =   1095
         Width           =   4440
      End
      Begin prjStudentIMS.ucHorizontal3DLine ucHorizontal3DLine1 
         Height          =   30
         Left            =   360
         TabIndex        =   3
         Top             =   4200
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   53
      End
      Begin VB.Label lblPath 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Currently selected  database location path "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   5
         Top             =   2640
         Width           =   3105
      End
      Begin VB.Label lblFileName 
         Appearance      =   0  'Flat
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   360
         TabIndex        =   4
         Top             =   2880
         Width           =   4440
      End
   End
End
Attribute VB_Name = "frmPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdOk_Click()
    frmBackup.txtPath.Text = lblFileName.Caption
    Call Unload(Me)
End Sub

Private Sub Dir1_Change()
    lblFileName.Caption = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error GoTo A1:
    Dir1.Path = Drive1.Drive
    Exit Sub
A1:
    MsgBox "The selected drive can not be accessed ...", vbInformation, "Drive not accessible ..."
    Drive1.Drive = "c:"
End Sub

Private Sub Form_Load()
    Drive1.Drive = "c:"
    lblFileName.Caption = Dir1.Path
    Call CenterForm(Me)
    Call initFrame(fraBackground)
End Sub

