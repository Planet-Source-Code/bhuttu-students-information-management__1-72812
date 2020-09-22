VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPatient 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patient"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6645
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
   ScaleHeight     =   6105
   ScaleWidth      =   6645
   StartUpPosition =   3  'Windows Default
   Begin prjStudentIMS.ucHorizontal3DLine ucHorizontal3DLine1 
      Height          =   30
      Left            =   -360
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   360
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   53
   End
   Begin prjStudentIMS.ucVertical3DLine ucVertical3DLine2 
      Height          =   285
      Left            =   2160
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   503
   End
   Begin prjStudentIMS.ucVertical3DLine ucVertical3DLine1 
      Height          =   255
      Left            =   1080
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjStudentIMS.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   375
      Left            =   0
      TabIndex        =   14
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
      TabIndex        =   15
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
      Left            =   6280
      TabIndex        =   17
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
      Height          =   5745
      Left            =   0
      TabIndex        =   22
      Top             =   360
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   10134
      IconSize        =   0
      CaptionColor    =   128
      Caption         =   "Existing Student Informations"
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
      Begin VB.TextBox txtClassID 
         Height          =   285
         Left            =   6120
         TabIndex        =   38
         Top             =   4800
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox txtAcadYRID 
         Height          =   285
         Left            =   6120
         TabIndex        =   37
         Top             =   2280
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton cmdCL 
         Height          =   285
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   4800
         Width           =   375
      End
      Begin VB.CommandButton cmdAY 
         Height          =   285
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox txtClass 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   4800
         Width           =   3255
      End
      Begin VB.TextBox txtAY 
         Height          =   285
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   2280
         Width           =   3255
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "Text2"
         Top             =   720
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   2280
         TabIndex        =   11
         Top             =   5160
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   503
         _Version        =   393216
         CustomFormat    =   "dd - MMM - yyyy"
         Format          =   52166659
         CurrentDate     =   38882
      End
      Begin VB.TextBox txtCPhone 
         Height          =   285
         Left            =   2280
         MaxLength       =   14
         TabIndex        =   4
         Top             =   2640
         Width           =   2535
      End
      Begin VB.TextBox txtRegdNo 
         Height          =   315
         Left            =   2280
         MaxLength       =   24
         TabIndex        =   9
         Top             =   4440
         Width           =   2565
      End
      Begin VB.TextBox txtEMail 
         Height          =   315
         Left            =   2280
         MaxLength       =   39
         TabIndex        =   6
         Top             =   3360
         Width           =   3765
      End
      Begin VB.TextBox txtMPhone 
         Height          =   315
         Left            =   2280
         MaxLength       =   14
         TabIndex        =   5
         Top             =   3000
         Width           =   2565
      End
      Begin VB.TextBox txtAddress 
         Height          =   765
         Left            =   2280
         MaxLength       =   249
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1440
         Width           =   3735
      End
      Begin VB.TextBox txtName 
         Height          =   315
         Left            =   2280
         MaxLength       =   49
         TabIndex        =   1
         Top             =   1080
         Width           =   3765
      End
      Begin VB.TextBox txtSection 
         Height          =   315
         Left            =   2280
         MaxLength       =   4
         TabIndex        =   7
         Top             =   3720
         Width           =   2565
      End
      Begin VB.TextBox txtRollNo 
         Height          =   315
         Left            =   2280
         MaxLength       =   9
         TabIndex        =   8
         Top             =   4080
         Width           =   2565
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student ID :"
         Height          =   195
         Left            =   600
         TabIndex        =   36
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Auto Generated)"
         Height          =   195
         Left            =   4920
         TabIndex        =   35
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Academic Year :"
         Height          =   255
         Left            =   600
         TabIndex        =   34
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class :"
         Height          =   195
         Left            =   600
         TabIndex        =   32
         Top             =   4800
         Width           =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Section :"
         Height          =   195
         Left            =   600
         TabIndex        =   31
         Top             =   3720
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail Address :"
         Height          =   195
         Left            =   600
         TabIndex        =   30
         Top             =   3360
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cell Phone No :"
         Height          =   195
         Left            =   600
         TabIndex        =   29
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Land Phone No :"
         Height          =   195
         Left            =   600
         TabIndex        =   28
         Top             =   2640
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Address :"
         Height          =   195
         Left            =   600
         TabIndex        =   27
         Top             =   1410
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Student Name :"
         Height          =   195
         Left            =   600
         TabIndex        =   26
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Roll No :"
         Height          =   195
         Left            =   600
         TabIndex        =   25
         Top             =   4080
         Width           =   600
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Registration No :"
         Height          =   195
         Left            =   600
         TabIndex        =   24
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Admission :"
         Height          =   195
         Left            =   600
         TabIndex        =   23
         Top             =   5160
         Width           =   1425
      End
   End
   Begin prjStudentIMS.ucVertical3DLine ucVertical3DLine3 
      Height          =   255
      Left            =   3330
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjStudentIMS.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   2160
      TabIndex        =   16
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
      TabIndex        =   18
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
Attribute VB_Name = "frmPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-- Global Variables.
Public g_lAddState  As Long
Dim txt


Private Sub init()

On Error GoTo errHandler
If g_lAddState = 1 Then
    Set rsStudent = New ADODB.Recordset
    rsStudent.Open "Select * FROM Students", g_cn, adOpenDynamic, adLockOptimistic
    rsStudent.AddNew
    Text2.Text = rsStudent!StudentID
    DTPicker1.Value = Date
End If

errHandler:
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
End Sub

Private Sub cmdAY_Click()
    frmSelectAcadYR.Show vbModal    '
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCL_Click()
    frmSelectClass.Show vbModal
End Sub

Private Sub cmdHelp_Click()
    MsgBox "The online help manual is not available with the demo version of this application.", vbExclamation, "Online Help Missing..."
End Sub

Private Sub cmdPrint_Click()
    MsgBox "The facility is not available with the current demo version of this program.", vbExclamation, "Demo Version Limitations..."
    Exit Sub
End Sub

Private Sub cmdSave_Click()
    Dim sName As String

    On Error GoTo errHandler

    '-- Required Field.
    If Not isValidData(txtName, "the name of the student. This field is mandatory") Then Exit Sub
    sName = Trim$(txtName)
    If Not isValidData(txtRollNo, "the roll number of the student. This field is mandatory") Then Exit Sub
    If Not isValidData(txtRegdNo, "the registration number of the student. This field is mandatory") Then Exit Sub
    If txtAY = "" Then
        MsgBox "Please select the academic year to which the student belongs to.", vbExclamation, "Incomplete Data..."
        txtAY.SetFocus
        Exit Sub
    ElseIf txtClass = "" Then
        MsgBox "Please select the class to which the student belongs to.", vbExclamation, "Incomplete Data..."
        txtClass.SetFocus
        Exit Sub
    End If
    
    
    '-- Save Data.
    
        If g_lAddState = 1 Then
        'With rsStudent
        'New data Addition
            rsStudent!StudName = sName
            rsStudent!Address = Trim$(txtAddress)
            rsStudent!AcadYearID = CLng(Trim$(txtAcadYRID))
            rsStudent!LPhone = Trim$(txtCPhone)
            rsStudent!CPhone = Trim$(txtMPhone)
            rsStudent!EMail = Trim$(txtEMail)
            rsStudent!Section = Trim$(txtSection)
            rsStudent!RollNo = Trim$(txtRollNo)
            rsStudent!RegdNo = Trim$(txtRegdNo)
            rsStudent!DOA = CDate(DTPicker1.Value)
            rsStudent!ClassID = CLng(Trim$(txtClassID))
            rsStudent.Update
        'End With
        rsStudent.Close
        Set rsStudent = Nothing
        ElseIf g_lAddState = 0 Then
            'Existing data editing
            Set rsStudent = New ADODB.Recordset
            rsStudent.Open "Select * FROM Students WHERE StudentID=" & CLng(Text2.Text) & "", g_cn, adOpenDynamic, adLockOptimistic
            With rsStudent
                .Fields("StudName") = sName
                .Fields("Address") = Trim$(txtAddress)
                .Fields("AcadYearID") = CLng(Trim$(txtAcadYRID))
                .Fields("LPhone") = Trim$(txtCPhone)
                .Fields("CPhone") = Trim$(txtMPhone)
                .Fields("EMail") = Trim$(txtEMail)
                .Fields("Section") = Trim$(txtSection)
                .Fields("RollNo") = Trim$(txtRollNo)
                .Fields("RegdNo") = Trim$(txtRegdNo)
                .Fields("DOA") = CDate(DTPicker1.Value)
                .Fields("ClassID") = CLng(Trim$(txtClassID))
                .Update
            End With
            rsStudent.Close
            Set rsStudent = Nothing
            MsgBox "The selected student's record has been successfully updated.", vbInformation, "Record Updated Successfully..."
        End If
       
    '-- Prompt user.
    If g_lAddState = 1 Then
        
        Dim lReply As Long
        
        Call frmMain.cmdRefresh_Click
        'Call PointToRecord(g_rsPatient, "StudentID", False, g_lOldID, 0)
        
        lReply = MsgBox("Do you want to add another new student record?", vbQuestion + vbYesNo, "Add New Record")
        If lReply = vbYes Then
            g_lAddState = 1
            Call ClearText
            Call init
            txtName.SetFocus
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
        cmdAY.Picture = .i16x16.ListImages(8).Picture
        cmdCL.Picture = .i16x16.ListImages(8).Picture
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
    Call initFrame(fraBackGround)
    
    Call init
    
    Call EndBusy
End Sub

Private Sub ClearText()
For Each txt In Me.Controls
    If TypeOf txt Is TextBox Then
        txt.Text = ""
    End If
Next
DTPicker1.Value = Date
End Sub
