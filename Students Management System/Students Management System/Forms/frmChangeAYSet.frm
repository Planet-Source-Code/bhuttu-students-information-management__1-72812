VERSION 5.00
Begin VB.Form frmChangeAYSet 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   ScaleHeight     =   2445
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjStudentIMS.ucGradContainer fraBackGround 
      Height          =   2450
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   4313
      IconSize        =   0
      Caption         =   "Modify the academic year setting"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdModify 
         Caption         =   "Modify"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancel 
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
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   1920
         Width           =   1215
      End
      Begin prjStudentIMS.ucHorizontal3DLine ucHorizontal3DLine1 
         Height          =   30
         Left            =   480
         TabIndex        =   5
         Top             =   1800
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   53
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "New academic year setting :"
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
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   2040
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Current academic year setting :"
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
         Left            =   480
         TabIndex        =   3
         Top             =   840
         Width           =   2280
      End
   End
End
Attribute VB_Name = "frmChangeAYSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub FillAcadYRName()
Combo1.AddItem "[Select Academic Year]"
Set rsTemp = New ADODB.Recordset
rsTemp.Open "Select AcadYR_Duration From AcadYR", g_cn, adOpenKeyset, adLockOptimistic
If Not rsTemp.EOF Or Not rsTemp.BOF Then
    Do While Not rsTemp.EOF
        Combo1.AddItem rsTemp!AcadYR_Duration
        rsTemp.MoveNext
    Loop
Else
    'do nothing
    'Resume Next
End If
'If rsAcct.State = adStateOpen Then rsAcct.Close
Set rsTemp = Nothing
Combo1.Text = "[Select Academic Year]"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdModify_Click()
On Error GoTo errorHandler
If user_level <> 1 Then
    MsgBox "You don't have the sufficient privilege to modify the academic year" & vbCrLf & _
           "setting. Only an administrative user can perform such a task.", vbExclamation, "Warning..."
    Exit Sub
Else
    If Combo1.Text = "[Select Academic Year]" Then
        MsgBox "Please select the new academic year you want to set for.", vbExclamation, "Select Academic Year..."
        Combo1.SetFocus
        Exit Sub
    Else
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open "Select * FROM Settings WHERE AcadYR_Setting ='" & Trim$(Text1.Text) & "'", g_cn, adOpenKeyset, adLockOptimistic
        rsTemp!AcadYR_Setting = CStr(Combo1.Text)
        rsTemp.Update
        rsTemp.Close
        Set rsTemp = Nothing
    
        '-- Close connection.
        g_cn.Close
    
        '-- Fresh DB Connection Again
        g_cn.Open g_sConnectionString
    
        '-- Reload Recordsets.
        frmMain.tvwReport.Nodes.Clear
        frmMain.Form_Load
        MsgBox "The application's academic year setting has been successfully changed to the new value.", vbInformation, "Academic Year Setting Changed..."
        Unload Me
        Unload frmOptions
    End If
End If
errorHandler:
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

Private Sub Form_Load()
    Call CenterForm(Me)
    Call initFrame(fraBackGround)
    Call FillAcadYRName
    '--- Global value
    Text1.Text = cur_acadYear
End Sub
