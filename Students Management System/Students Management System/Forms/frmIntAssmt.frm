VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIntAssmt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6645
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin prjStudentIMS.ucHorizontal3DLine ucHorizontal3DLine1 
      Height          =   30
      Left            =   -360
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   360
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   53
   End
   Begin prjStudentIMS.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Save"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjStudentIMS.lvButtons_H cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Cancel"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjStudentIMS.lvButtons_H cmdHelp 
      Height          =   375
      Left            =   6285
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjStudentIMS.ucGradContainer fraBackground 
      Height          =   5385
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   9499
      IconSize        =   0
      CaptionColor    =   128
      Caption         =   "Assignments For Students"
      CaptionAlignment=   2
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   2400
         TabIndex        =   31
         Top             =   3600
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd - MMM - yyyy"
         Format          =   20709379
         CurrentDate     =   38886
      End
      Begin VB.TextBox txtMarks 
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
         Left            =   2400
         TabIndex        =   28
         Top             =   2880
         Width           =   2535
      End
      Begin VB.TextBox txtRoll 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtSection 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtClass 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtTerm 
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
         Left            =   2400
         MaxLength       =   49
         TabIndex        =   20
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000012&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   12
         Text            =   "Text2"
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtRemarks 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1035
         Left            =   2400
         MaxLength       =   99
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   3960
         Width           =   3765
      End
      Begin VB.CommandButton cmdAT 
         Height          =   285
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtStudent 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CommandButton cmdAB 
         Height          =   285
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3240
         Width           =   375
      End
      Begin VB.TextBox txtStaff 
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
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3240
         Width           =   3255
      End
      Begin VB.TextBox txtStaffID 
         Height          =   285
         Left            =   6240
         TabIndex        =   6
         Top             =   3240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtStudID 
         Height          =   285
         Left            =   6240
         TabIndex        =   5
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of assessment :"
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
         Left            =   600
         TabIndex        =   30
         Top             =   3600
         Width           =   1530
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Marks secured :"
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
         Left            =   600
         TabIndex        =   29
         Top             =   2880
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Roll No :"
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
         Left            =   600
         TabIndex        =   27
         Top             =   2520
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section :"
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
         Left            =   600
         TabIndex        =   26
         Top             =   2160
         Width           =   630
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class :"
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
         Left            =   600
         TabIndex        =   25
         Top             =   1800
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assessment term :"
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
         Left            =   600
         TabIndex        =   21
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assessment ID :"
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
         Left            =   600
         TabIndex        =   17
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Auto Generated)"
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
         Left            =   5040
         TabIndex        =   16
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "For student :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   15
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assessment done by :"
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
         Left            =   600
         TabIndex        =   14
         Top             =   3240
         Width           =   1590
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks :"
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
         Left            =   600
         TabIndex        =   13
         Top             =   3960
         Width           =   720
      End
   End
   Begin prjStudentIMS.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   2160
      TabIndex        =   18
      Top             =   0
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "Print"
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Focus           =   0   'False
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjStudentIMS.lvButtons_H cmdHeader 
      Height          =   375
      Left            =   -30
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   661
      CapAlign        =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      Enabled         =   0   'False
      cBack           =   -2147483633
   End
End
Attribute VB_Name = "frmIntAssmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-- Global Variables.
Public g_AssmtAddState  As Long
Dim txt


Private Sub init()

On Error GoTo errHandler
If g_AssmtAddState = 1 Then
    Set rsStudent = New ADODB.Recordset
    rsStudent.Open "Select * FROM IntAssmt", g_cn, adOpenDynamic, adLockOptimistic
    rsStudent.AddNew
    Text2.Text = rsStudent!IntAssmtID
End If

errHandler:
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
End Sub

Private Sub cmdAB_Click()
    reqdFor_BJrTrans = False
    reqdFor_IntAssmt = True
    reqdFor_Assignmt = False
    frmSelectStaff.Show vbModal
End Sub

Private Sub cmdAT_Click()
    reqdFor_BJrTrans = False
    reqdFor_IntAssmt = True
    reqdFor_Assignmt = False
    frmSelectStud.Show vbModal
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    MsgBox "The online help manual is not available with the demo version of this application.", vbExclamation, "Online Help Missing..."
End Sub

Private Sub cmdPrint_Click()
    MsgBox "The facility is not available with the current demo version of this program.", vbExclamation, "Demo Version Limitations..."
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    Dim sDescp As String

   On Error GoTo errHandler

    If Not isValidData(txtTerm, "a suitable term name for the internal assessment being carried out") Then Exit Sub
    If Not isValidData(txtStudent, "or select the name of the student whose internal assessment mark is to be entered") Then Exit Sub
    If Not isValidData(txtStaff, "or select the name of the teacher by whom the assessment is being done") Then Exit Sub
    If Not isValidData(txtRemarks, "a suitable remark for the internal assessment being done") Then Exit Sub
    If Not IsNumeric(txtMarks) Then
        MsgBox "The internal assessment mark entered is not a valid numeric data.", vbExclamation, "Invalid Data..."
        txtMarks.Text = ""
        txtMarks.SetFocus
        Exit Sub
    ElseIf Val(txtMarks) > 100 Then
        MsgBox "The internal assessment mark is too unrealistic and high.", vbExclamation, "Invalid Mark..."
        txtMarks.Text = ""
        txtMarks.SetFocus
        Exit Sub
    End If
    '-- Save Data.
    If g_AssmtAddState = 1 Then
        With rsStudent
        'New data Addition
            .Fields("Remarks") = Trim$(txtRemarks)
            .Fields("DOA") = CDate(DTPicker1.Value)
            .Fields("IntAssmt_Term") = Trim$(txtTerm)
            .Fields("StaffID") = CLng(txtStaffID)
            .Fields("StudentID") = CLng(txtStudID)
            .Fields("IntAssmt_Marks") = CInt(txtMarks)
            .Update
        End With
        rsStudent.Close
        Set rsStudent = Nothing
        ElseIf g_AssmtAddState = 0 Then
            'Existing data editing
            Set rsStudent = New ADODB.Recordset
            rsStudent.Open "Select * FROM IntAssmt WHERE IntAssmtID=" & CLng(Text2.Text) & "", g_cn, adOpenDynamic, adLockOptimistic
            With rsStudent
                .Fields("Remarks") = Trim$(txtRemarks)
                .Fields("DOA") = CDate(DTPicker1.Value)
                .Fields("IntAssmt_Term") = Trim$(txtTerm)
                .Fields("StaffID") = CLng(txtStaffID)
                .Fields("StudentID") = CLng(txtStudID)
                .Fields("IntAssmt_Marks") = CInt(txtMarks)
                .Update
            End With
            rsStudent.Close
            Set rsStudent = Nothing
            MsgBox "The selected student's internal assessment mark has been syccessfully updated.", vbInformation, "Record Updated Successfully..."
        End If
        
    '-- Prompt user.
    If g_AssmtAddState = 1 Then
        
        Dim lReply As Long
        
        Call frmMain.cmdRefresh_Click
        'Call PointToRecord(g_rsPatient, "StudentID", False, g_lOldID, 0)
        
        lReply = MsgBox("Do you want to add another new internal assessment record?", vbQuestion + vbYesNo, "Add Another Record")
        If lReply = vbYes Then
            g_AssmtAddState = 1
            Call ClearText
            Call init
            txtTerm.SetFocus
        Else
            Unload Me
        End If
    Else
        Call frmMain.cmdRefresh_Click
        'Call PointToRecord(g_rsPatient, "StudentID", False, g_lOldID, 0)
        Unload Me
    End If

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmPatient", "cmdSave_Click", True)
    End If
End Sub

Private Sub Form_Load()
    Call StartBusy
    
    'txtDate.Text = FormatDateTime(Date, vbLongDate)

    With frmMain
        cmdAT.Picture = .i16x16.ListImages(8).Picture
        cmdAB.Picture = .i16x16.ListImages(8).Picture
    End With
    
    Icon = frmMain.Icon
    Call CenterForm(Me)
    
    Call initToolbar(cmdHeader, frmMain.imlTrans)
    Call initToolbar(cmdSave, frmMain.imlTrans, , g_eIcon.Save)
    Call initToolbar(cmdCancel, frmMain.imlTrans, , g_eIcon.Cancel)
    Call initToolbar(cmdPrint, frmMain.imlTrans, , g_eIcon.Print)
    Call initToolbar(cmdHelp, frmMain.imlTrans, , g_eIcon.Help)
    'Call initFrame(fraBackground, frmMain.imlPanel, g_eIcon.Patient)
    'Original Line above for which an icon appears
    Call initFrame(fraBackground)
    
    Call init
    
    Call EndBusy
End Sub

Private Sub ClearText()
For Each txt In Me.Controls
    If TypeOf txt Is TextBox Then
        txt.Text = ""
    End If
Next
End Sub

