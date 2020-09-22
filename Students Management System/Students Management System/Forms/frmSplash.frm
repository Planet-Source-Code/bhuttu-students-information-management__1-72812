VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   5895
   ScaleWidth      =   6225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   2880
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   960
      Picture         =   "frmSplash.frx":4DC3
      Top             =   4080
      Width           =   1260
   End
   Begin VB.Shape Shape1 
      Height          =   5895
      Left            =   0
      Top             =   0
      Width           =   6225
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by: Partha S. Paul"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   1
      Top             =   5520
      Width           =   2760
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Student IMS [Demo Version]"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   240
      TabIndex        =   0
      Top             =   5160
      Width           =   3540
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim abc As Integer

Private Sub Form_Load()
    'lblVersion = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Label1.ForeColor = vbBlack
    Label2.ForeColor = vbBlue
    Call CenterForm(Me)
    Timer1.Enabled = True
    Timer1.Interval = 50
End Sub

Private Sub Image1_Click()
    frmLogin.Show
    Call Unload(Me)
End Sub

Private Sub Timer1_Timer()
    If abc = 300 Then
        Call Image1_Click
    Else
        abc = abc + 30
    End If
    
End Sub
