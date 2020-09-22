VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBookJrTrans 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
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
      Height          =   4545
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   8017
      IconSize        =   0
      CaptionColor    =   128
      Caption         =   "Book/Reference Transection Record"
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
      Begin VB.TextBox txtAuthors 
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
         TabIndex        =   30
         Top             =   1800
         Width           =   3735
      End
      Begin VB.ComboBox cboRetrn 
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
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3960
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtReturn 
         Height          =   285
         Left            =   2400
         TabIndex        =   27
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
      Begin VB.TextBox txtBookID 
         Height          =   285
         Left            =   6240
         TabIndex        =   15
         Top             =   1440
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtStudID 
         Height          =   285
         Left            =   6240
         TabIndex        =   14
         Top             =   2160
         Visible         =   0   'False
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
         TabIndex        =   13
         Top             =   2160
         Width           =   3255
      End
      Begin VB.CommandButton cmdAB 
         Height          =   285
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2160
         Width           =   375
      End
      Begin VB.TextBox txtTitle 
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
         TabIndex        =   11
         Top             =   1440
         Width           =   3255
      End
      Begin VB.CommandButton cmdAT 
         Height          =   285
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1440
         Width           =   375
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
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   720
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
         TabIndex        =   8
         Top             =   3240
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
         TabIndex        =   7
         Top             =   2880
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
         TabIndex        =   6
         Top             =   2520
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker dtIssue 
         Height          =   285
         Left            =   2400
         TabIndex        =   5
         Top             =   1080
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Author(s) :"
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
         TabIndex        =   31
         Top             =   1800
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Has returned ?"
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
         Top             =   3960
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of return :"
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
         TabIndex        =   24
         Top             =   3600
         Width           =   1140
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Issued to :"
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
         TabIndex        =   23
         Top             =   2160
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Boook/Ref title :"
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
         TabIndex        =   22
         Top             =   1440
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
         TabIndex        =   21
         Top             =   720
         Width           =   1275
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Transection ID :"
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
         TabIndex        =   20
         Top             =   720
         Width           =   1155
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
         TabIndex        =   19
         Top             =   3240
         Width           =   630
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
         TabIndex        =   18
         Top             =   2880
         Width           =   600
      End
      Begin VB.Label Label9 
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
         TabIndex        =   17
         Top             =   2520
         Width           =   480
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date of issue :"
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
         TabIndex        =   16
         Top             =   1080
         Width           =   1050
      End
   End
   Begin prjStudentIMS.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   2160
      TabIndex        =   25
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
      TabIndex        =   26
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
Attribute VB_Name = "frmBookJrTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-- Global Variables.
Public g_BJrAddState  As Long

Dim txt


Private Sub init()

On Error GoTo errHandler
If g_BJrAddState = 1 Then
    Set rsStudent = New ADODB.Recordset
    rsStudent.Open "Select * FROM BJTransection", g_cn, adOpenDynamic, adLockOptimistic
    rsStudent.AddNew
    Text2.Text = rsStudent!TransID
    dtIssue = CDate(Date)
    dtReturn = CDate(Date + 7)
    cboRetrn.Text = "To be returned later"
    cboRetrn.Enabled = False
End If

errHandler:
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
End Sub

Private Sub cmdAB_Click()
    reqdFor_BJrTrans = True
    reqdFor_IntAssmt = False
    reqdFor_Assignmt = False
    frmSelectStud.Show vbModal
End Sub

Private Sub cmdAT_Click()
    reqdFor_BJrTrans = True
    frmSelectBookJr.Show vbModal
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
    Dim copies_issued As Integer

    On Error GoTo errHandler

    If Not isValidData(txtTitle, "or select the title of the book/reference to be issued") Then Exit Sub
    If Not isValidData(txtStudent, "or select the name of the student to whom the book/reference is to be issued") Then Exit Sub
    
    'If (Abs(CDate(dtIssue.Value) - Date) <> 0) Then
    '    MsgBox "Wrong date"
    'ElseIf (Abs(CDate(dtIssue.Value) - CDate(dtReturn.Value)) <> 7) Then
    '    MsgBox "Wrong"
    'End If
    
    
    '-- Save Data.
    If g_BJrAddState = 1 Then
        If Abs(Date - CDate(dtIssue.Value)) <> 0 Then
            MsgBox "Date of issue cannot be earlier or later than the current date.", vbExclamation, "Invalid Data..."
            Exit Sub
        ElseIf (Abs(CDate(dtReturn.Value) - CDate(dtIssue.Value)) > 7) Then
            MsgBox "Date of return cannot be more than the 7 days from the date of issue.", vbExclamation, "Invalid Data..."
            Exit Sub
        End If
        
        Set rsTemp = New ADODB.Recordset
        rsTemp.Open "Select * FROM BJTransection WHERE BookJrID=" & CLng(txtBookID) & " AND StudentID=" & CLng(txtStudID) & " AND Returned = 0", g_cn, adOpenKeyset, adLockOptimistic
        If rsTemp.BOF Or rsTemp.EOF Then
        
        '--------------------------------------------------
        With rsStudent
        'New data Addition
            .Fields("DOI") = CDate(dtIssue.Value)
            .Fields("DOR") = CDate(dtReturn.Value)
            .Fields("BookJrID") = CLng(txtBookID)
            .Fields("StudentID") = CLng(txtStudID)
            .Fields("Returned") = 0
            .Update
        End With
        If rsStudent.State = adStateOpen Then rsStudent.Close
        rsStudent.Open "Select * FROM BookJournel WHERE BookJrID=" & CLng(Trim$(txtBookID.Text)) & "", g_cn, adOpenKeyset, adLockOptimistic
        'find the number of copies already issued
        copies_issued = CInt(rsStudent!CopiesIssued)
        'increase the number by 1
        copies_issued = copies_issued + 1
        'Update the record
        rsStudent!CopiesIssued = CInt(copies_issued)
        rsStudent.Update
        rsStudent.Close
        Set rsStudent = Nothing
        '--------------------------------------------------------
        Else
            MsgBox "It seems that a copy of " & CStr(txtTitle) & " is already lying with " & CStr(txtStudent) & ".", vbExclamation, "Already Borrowed..."
            Set rsTemp = Nothing
            Exit Sub
        End If
        '---------------------------------------------------------
    ElseIf g_BJrAddState = 0 Then
        'Existing data editing
        Set rsStudent = New ADODB.Recordset
        rsStudent.Open "Select * FROM BJTransection WHERE TransID=" & CLng(Text2.Text) & "", g_cn, adOpenDynamic, adLockOptimistic
        With rsStudent
            '.Fields("DOI") = CDate(dtIssue.Value)
            .Fields("DOR") = CDate(Date)
            '.Fields("BookJrID") = CLng(txtBookID)
            '.Fields("StudentID") = CLng(txtStudID)
            .Fields("Returned") = 1
            .Update
        End With
        If rsStudent.State = adStateOpen Then rsStudent.Close
        rsStudent.Open "Select * FROM BookJournel WHERE BookJrID=" & CLng(Trim$(txtBookID.Text)) & "", g_cn, adOpenKeyset, adLockOptimistic
        'find the number of copies already issued
        copies_issued = CInt(rsStudent!CopiesIssued)
        'reduce the number by 1 as book is now returned
        copies_issued = copies_issued - 1
        'Update the record
        rsStudent!CopiesIssued = CInt(copies_issued)
        rsStudent.Update
        rsStudent.Close
        Set rsStudent = Nothing
        MsgBox "The book/reference material data has been updated successfully.", vbInformation, "Record Updated Successfully..."
    End If
        
    '-- Prompt user.
    If g_BJrAddState = 1 Then
        
        Dim lReply As Long
        
        Call frmMain.cmdRefresh_Click
        'Call PointToRecord(g_rsPatient, "StudentID", False, g_lOldID, 0)
        
        lReply = MsgBox("Do you want to add another book/reference transection record?", vbQuestion + vbYesNo, "Add Another Record")
        If lReply = vbYes Then
            g_BJrAddState = 1
            Call ClearText
            Call init
            txtTitle.SetFocus
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
    
    cboRetrn.AddItem "To be returned later"
    'cboRetrn.AddItem "Returned back on time"
    cboRetrn.AddItem "Returning now..."
    
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
End Sub


