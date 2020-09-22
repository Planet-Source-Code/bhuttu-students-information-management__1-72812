VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmStaff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
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
      Height          =   6375
      Left            =   0
      TabIndex        =   4
      Top             =   360
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   11245
      IconSize        =   0
      CaptionColor    =   128
      Caption         =   "Departmental Staff Informations"
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
      Begin VB.ComboBox cboStatus 
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
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   5880
         Width           =   2655
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
         Height          =   735
         Left            =   2280
         MaxLength       =   99
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   5040
         Width           =   3735
      End
      Begin VB.TextBox txtQual 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2280
         MaxLength       =   149
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   4320
         Width           =   3765
      End
      Begin VB.TextBox txtDesignation 
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
         Left            =   2280
         MaxLength       =   39
         TabIndex        =   13
         Top             =   3240
         Width           =   2565
      End
      Begin VB.TextBox txtName 
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
         Left            =   2280
         MaxLength       =   49
         TabIndex        =   12
         Top             =   960
         Width           =   3765
      End
      Begin VB.TextBox txtAddress 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2280
         MaxLength       =   249
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1320
         Width           =   3735
      End
      Begin VB.TextBox txtMPhone 
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
         Left            =   2280
         MaxLength       =   15
         TabIndex        =   10
         Top             =   2520
         Width           =   2565
      End
      Begin VB.TextBox txtEMail 
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
         Left            =   2280
         TabIndex        =   9
         Top             =   2880
         Width           =   3765
      End
      Begin VB.TextBox txtSpecial 
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
         Left            =   2280
         MaxLength       =   49
         TabIndex        =   8
         Top             =   3960
         Width           =   3765
      End
      Begin VB.TextBox txtCPhone 
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
         Left            =   2280
         MaxLength       =   12
         TabIndex        =   7
         Top             =   2160
         Width           =   2535
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
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   600
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   2280
         TabIndex        =   6
         Top             =   3600
         Width           =   2655
         _ExtentX        =   4683
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
         Format          =   17629187
         CurrentDate     =   38882
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Staff status :"
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
         Top             =   5880
         Width           =   960
      End
      Begin VB.Label Label6 
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
         TabIndex        =   29
         Top             =   5040
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Joining :"
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
         Top             =   3600
         Width           =   1215
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Specialisation :"
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
         Top             =   3960
         Width           =   1065
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qualification :"
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
         Top             =   4320
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Staff Name :"
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
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contact Address :"
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
         Top             =   1290
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Land Phone No :"
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
         Top             =   2160
         Width           =   1185
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cell Phone No :"
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
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail Address :"
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
         Width           =   1155
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Designation :"
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
         Top             =   3240
         Width           =   945
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
         Left            =   4920
         TabIndex        =   16
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Staff ID :"
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
         TabIndex        =   15
         Top             =   600
         Width           =   675
      End
   End
   Begin prjStudentIMS.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   2160
      TabIndex        =   26
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
      cGradient       =   0
      Mode            =   0
      Value           =   0   'False
      cBack           =   -2147483633
   End
   Begin prjStudentIMS.lvButtons_H cmdHeader 
      Height          =   375
      Left            =   -30
      TabIndex        =   27
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
Attribute VB_Name = "frmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-- Global Variables.
Public g_StaffAddState  As Long
Dim txt


Public Sub init()

On Error GoTo errHandler
If g_StaffAddState = 1 Then
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "Select * FROM Staff", g_cn, adOpenDynamic, adLockOptimistic
    rsTemp.AddNew
    Text2.Text = rsTemp!StaffID
    DTPicker1.Value = Date
End If

errHandler:
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
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
    Dim sName As String

    On Error GoTo errHandler

    '-- Required Field.
    If Not isValidData(txtName, "the name of the staff. This field is mandatory") Then Exit Sub
    sName = Trim$(txtName)
    If Not isValidData(txtDesignation, "the designation of the staff. This field is mandatory") Then Exit Sub
    If Not isValidData(txtSpecial, "the field of specialisation of the staff. This field is mandatory") Then Exit Sub
    If Not isValidData(txtQual, "the educational qualification of the staff. This field is mandatory") Then Exit Sub
    If Not isValidData(cboStatus, "status of the staff. Staff status selection is mandatory") Then Exit Sub
    
    '-- Save Data.
    
        If g_StaffAddState = 1 Then
        With rsTemp
        'New data Addition
            .Fields("StaffName") = sName
            .Fields("Address") = Trim$(txtAddress)
            .Fields("Designation") = Trim$(txtDesignation)
            .Fields("Phone_No") = Trim$(txtCPhone)
            .Fields("MPhone_No") = Trim$(txtMPhone)
            .Fields("EMail") = Trim$(txtEMail)
            .Fields("DOJ") = CDate(DTPicker1.Value)
            .Fields("Specialisation") = Trim$(txtSpecial)
            .Fields("Qualification") = Trim$(txtQual)
            .Fields("Remarks") = Trim$(txtRemarks)
            .Fields("Status") = Trim$(cboStatus.Text)
            .Update
        End With
        rsTemp.Close
        Set rsTemp = Nothing
        ElseIf g_StaffAddState = 0 Then
            'Existing data editing
            Set rsTemp = New ADODB.Recordset
            rsTemp.Open "Select * FROM Staff WHERE StaffID=" & CLng(Text2.Text) & "", g_cn, adOpenDynamic, adLockOptimistic
            With rsTemp
                .Fields("StaffName") = sName
                .Fields("Address") = Trim$(txtAddress)
                .Fields("Designation") = Trim$(txtDesignation)
                .Fields("Phone_No") = Trim$(txtCPhone)
                .Fields("MPhone_No") = Trim$(txtMPhone)
                .Fields("EMail") = Trim$(txtEMail)
                .Fields("DOJ") = CDate(DTPicker1.Value)
                .Fields("Specialisation") = Trim$(txtSpecial)
                .Fields("Qualification") = Trim$(txtQual)
                .Fields("Remarks") = Trim$(txtRemarks)
                .Fields("Status") = Trim$(cboStatus.Text)
                .Update
            End With
            rsTemp.Close
            Set rsTemp = Nothing
            MsgBox "The staff record has been successfully updated to the database.", vbInformation, "Record Updated Successfully..."
        End If
        
    '-- Prompt user.
    If g_StaffAddState = 1 Then
        
        Dim lReply As Long
        
        Call frmMain.cmdRefresh_Click
        'Call PointToRecord(g_rsPatient, "StudentID", False, g_lOldID, 0)
        
        lReply = MsgBox("Do you want to add another new staff record?", vbQuestion + vbYesNo, "Add New Record")
        If lReply = vbYes Then
            g_StaffAddState = 1
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
    
    Icon = frmMain.Icon
    Call CenterForm(Me)
    
    cboStatus.AddItem "Membership Active"
    cboStatus.AddItem "Membership Expired"
    
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
cboStatus.Text = "Membership Active"
End Sub

