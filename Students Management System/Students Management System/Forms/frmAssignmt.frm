VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAssignmt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assignments"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6660
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
   ScaleHeight     =   4425
   ScaleWidth      =   6660
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
      Height          =   4065
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   7170
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
      Begin VB.TextBox txtStudID 
         Height          =   285
         Left            =   6240
         TabIndex        =   23
         Top             =   1800
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtStaffID 
         Height          =   285
         Left            =   6240
         TabIndex        =   22
         Top             =   2160
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAB 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2160
         Width           =   3255
      End
      Begin VB.CommandButton cmdAB 
         Height          =   285
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtAT 
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1800
         Width           =   3255
      End
      Begin VB.CommandButton cmdAT 
         Height          =   285
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1800
         Width           =   375
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   285
         Left            =   2400
         TabIndex        =   17
         Top             =   1440
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd - MMM - yyyy"
         Format          =   20709379
         CurrentDate     =   38884
      End
      Begin VB.TextBox txtDescp 
         Height          =   1275
         Left            =   2400
         MaxLength       =   249
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2520
         Width           =   3765
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
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   720
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   2400
         TabIndex        =   6
         Top             =   1080
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd - MMM - yyyy"
         Format          =   20709379
         CurrentDate     =   38882
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assignment dated :"
         Height          =   195
         Left            =   600
         TabIndex        =   14
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To be submitted by :"
         Height          =   195
         Left            =   600
         TabIndex        =   13
         Top             =   1440
         Width           =   1485
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assignment description :"
         Height          =   195
         Left            =   600
         TabIndex        =   12
         Top             =   2520
         Width           =   1755
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assigned by :"
         Height          =   195
         Left            =   600
         TabIndex        =   11
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Assignment to :"
         Height          =   255
         Left            =   600
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Auto Generated)"
         Height          =   195
         Left            =   5040
         TabIndex        =   9
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Assignment ID :"
         Height          =   195
         Left            =   600
         TabIndex        =   8
         Top             =   720
         Width           =   1140
      End
   End
   Begin prjStudentIMS.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   2160
      TabIndex        =   15
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
      TabIndex        =   16
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
Attribute VB_Name = "frmAssignmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-- Global Variables.
Public g_AssignAddState  As Long
Dim txt


Private Sub init()

On Error GoTo errHandler
If g_AssignAddState = 1 Then
    Set rsStudent = New ADODB.Recordset
    rsStudent.Open "Select * FROM Assignments", g_cn, adOpenDynamic, adLockOptimistic
    rsStudent.AddNew
    Text2.Text = rsStudent!AssignID
End If

errHandler:
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
End Sub

Private Sub cmdAB_Click()
    reqdFor_BJrTrans = False
    reqdFor_IntAssmt = False
    reqdFor_Assignmt = True
    frmSelectStaff.Show vbModal
End Sub

Private Sub cmdAT_Click()
    reqdFor_BJrTrans = False
    reqdFor_IntAssmt = False
    reqdFor_Assignmt = True
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

    '-- Required Field.
    If Not isValidData(txtAT, "or select the name of the student to whom the task is to be assigned") Then Exit Sub
    If Not isValidData(txtAB, "or select the name of the teacher by whom the task is assigned to the student") Then Exit Sub
    If Not isValidData(txtDescp, "a description of the assignment given to the student for carrying out") Then Exit Sub
    sDescp = Trim$(txtDescp)
    
    '-- Save Data.
    
        If g_AssignAddState = 1 Then
        With rsStudent
        'New data Addition
            .Fields("Description") = sDescp
            .Fields("DOA") = CDate(DTPicker1.Value)
            .Fields("DOS") = CDate(DTPicker2.Value)
            .Fields("StudentID") = CLng(txtStudID)
            .Fields("StaffID") = CLng(txtStaffID)
            .Update
        End With
        rsStudent.Close
        Set rsStudent = Nothing
        ElseIf g_AssignAddState = 0 Then
            'Existing data editing
            Set rsStudent = New ADODB.Recordset
            rsStudent.Open "Select * FROM Assignments WHERE AssignID=" & CLng(Text2.Text) & "", g_cn, adOpenDynamic, adLockOptimistic
            With rsStudent
                .Fields("Description") = sDescp
                .Fields("DOA") = CDate(DTPicker1.Value)
                .Fields("DOS") = CDate(DTPicker2.Value)
                .Fields("StudentID") = CLng(txtStudID)
                .Fields("StaffID") = CLng(txtStaffID)
                .Update
            End With
            rsStudent.Close
            Set rsStudent = Nothing
            MsgBox "Assignment Record has successfully been updated.", vbInformation, "Assignment Record Updated..."
        End If
        
    '-- Prompt user.
    If g_AssignAddState = 1 Then
        
        Dim lReply As Long
        
        Call frmMain.cmdRefresh_Click
        'Call PointToRecord(g_rsPatient, "StudentID", False, g_lOldID, 0)
        
        lReply = MsgBox("Do you want to add another new assignment record?", vbQuestion + vbYesNo, "Add New Assignment")
        If lReply = vbYes Then
            g_AssignAddState = 1
            Call ClearText
            Call init
            txtDescp.SetFocus
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
    
    DTPicker1.Value = Date
    DTPicker2.Value = Date + 7
    
    Call EndBusy
End Sub

Private Sub ClearText()
For Each txt In Me.Controls
    If TypeOf txt Is TextBox Then
        txt.Text = ""
    End If
Next
DTPicker1.Value = Date
DTPicker2.Value = Date + 7
End Sub
