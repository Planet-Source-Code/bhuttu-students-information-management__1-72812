VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   0  'None
   Caption         =   "Options"
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6255
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjStudentIMS.ucGradContainer fraBackGround 
      Height          =   5340
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   9419
      IconSize        =   0
      Caption         =   "Application Settings"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4680
         TabIndex        =   3
         Top             =   4830
         Width           =   1215
      End
      Begin prjStudentIMS.ucHorizontal3DLine ucHorizontal3DLine1 
         Height          =   30
         Left            =   600
         TabIndex        =   2
         Top             =   4680
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   53
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   840
         MouseIcon       =   "frmOptions.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmOptions.frx":0152
         Top             =   3960
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   840
         MouseIcon       =   "frmOptions.frx":06F4
         MousePointer    =   99  'Custom
         Picture         =   "frmOptions.frx":0846
         Top             =   3160
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   840
         MouseIcon       =   "frmOptions.frx":0DD2
         MousePointer    =   99  'Custom
         Picture         =   "frmOptions.frx":0F24
         Top             =   2360
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   840
         MouseIcon       =   "frmOptions.frx":14DB
         MousePointer    =   99  'Custom
         Picture         =   "frmOptions.frx":162D
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add or edit application user and their access level"
         Height          =   195
         Left            =   1680
         TabIndex        =   7
         Top             =   4080
         Width           =   3555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Change the academic year setting"
         Height          =   195
         Left            =   1680
         TabIndex        =   6
         Top             =   3280
         Width           =   2460
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "View user log and examine the usage statistics"
         Height          =   195
         Left            =   1680
         TabIndex        =   5
         Top             =   2475
         Width           =   3345
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Lock the application if left unattended for some time"
         Height          =   195
         Left            =   1680
         TabIndex        =   4
         Top             =   1680
         Width           =   3735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmOptions.frx":1C10
         Height          =   615
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Width           =   5415
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub Image1_Click(Index As Integer)
Select Case Index
    Case 0, 1, 3
        MsgBox "This setting option is not available with the demo version of this application.", vbExclamation, "Demo Version Limitations..."
        Exit Sub
    Case 2
        Load frmChangeAYSet
        frmChangeAYSet.Show vbModal
End Select
End Sub

Private Sub Form_Load()
    
    Call CenterForm(Me)
    Call initFrame(fraBackGround)
    'With frmMain
    '    Command1(0).Picture = .imlPanel.ListImages(44).Picture
    '    Command1(1).Picture = .imlPanel.ListImages(20).Picture
    '    Command1(2).Picture = .imlPanel.ListImages(31).Picture
    '    Command1(3).Picture = .imlPanel.ListImages(19).Picture
    'End With
    
End Sub
