VERSION 5.00
Begin VB.Form frmClass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fee"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   6465
   StartUpPosition =   3  'Windows Default
   Begin prjStudentIMS.ucHorizontal3DLine ucHorizontal3DLine1 
      Height          =   30
      Left            =   -360
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   53
   End
   Begin prjStudentIMS.ucGradContainer fraBackground 
      Height          =   2325
      Left            =   15
      TabIndex        =   6
      Top             =   360
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   4101
      IconSize        =   0
      CaptionColor    =   128
      Caption         =   "Add/Edit Class Records"
      CaptionAlignment=   2
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtClassName 
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
         Left            =   2040
         MaxLength       =   49
         TabIndex        =   1
         Top             =   960
         Width           =   3945
      End
      Begin VB.TextBox txtClassID 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtDescp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2040
         MaxLength       =   99
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1320
         Width           =   3945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "(Auto generated ID #)"
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
         Left            =   4320
         TabIndex        =   14
         Top             =   600
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class ID # :"
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
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name of the class :"
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
         TabIndex        =   12
         Top             =   960
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Brief description :"
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
         TabIndex        =   7
         Top             =   1320
         Width           =   1260
      End
   End
   Begin prjStudentIMS.ucVertical3DLine ucVertical3DLine2 
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   503
   End
   Begin prjStudentIMS.ucVertical3DLine ucVertical3DLine1 
      Height          =   255
      Left            =   1080
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   60
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   450
   End
   Begin prjStudentIMS.lvButtons_H cmdSave 
      Default         =   -1  'True
      Height          =   375
      Left            =   30
      TabIndex        =   3
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
      Height          =   375
      Left            =   1080
      TabIndex        =   4
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
      Left            =   6120
      TabIndex        =   10
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
   Begin prjStudentIMS.lvButtons_H cmdHeader 
      Height          =   375
      Left            =   -30
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   6555
      _ExtentX        =   11562
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
Attribute VB_Name = "frmClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public g_AddClassState  As Long
'Description
'Something

Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub cmdHelp_Click()
    MsgBox "The online help manual is not available with the demo version of this application.", vbExclamation, "Online Help Missing..."
End Sub

Private Sub cmdSave_Click()

    On Error GoTo errHandler


    '-- Required Field.
    If Not isValidData(txtClassName, "the name class to be added to the database") Then Exit Sub
    If Not isValidData(txtDescp, "a brief description of the class to be added to the database") Then Exit Sub
    
    
    '-- Check for Unique Class Name.
    'If g_AddClassState = 1 Then
    '    If isExisting(g_cn, "Class", "ClassName='" & Trim$(CStr(txtClassName)) & "'") Then
    '        MsgBox "The class name as entered already exists in the database.", vbExclamation, "Duplicate Data Error..."
    '        Exit Sub
    '    End If
    'Else
    '    '-- Check if user updated the class table.
    '    If InStr(1, sName, g_sOldName) = 0 Then
    '        If isExisting(g_cn, "tblPatient", "Name='" & sName & "' AND ID<>" & g_lOldID) Then
    '            MsgBox "Patient Name already exists.", vbExclamation
    '            Exit Sub
    '        End If
    '    End If
    'End If
    
    
    '-- Save Data.
    If g_AddClassState = 1 Then
        With rsStudent
            .Fields("ClassName") = Trim$(txtClassName.Text)
            .Fields("Description") = Trim$(txtDescp)
            .Update
        End With
        Call frmSelectClass.reload_rec
        rsStudent.Close
        Set rsStudent = Nothing
        
    ElseIf g_AddClassState = 0 Then
        Set rsStudent = New ADODB.Recordset
        rsStudent.Open "SELECT * FROM Class WHERE ClassID=" & CLng(txtClassID.Text) & "", g_cn, adOpenKeyset, adLockOptimistic
        With rsStudent
            .Fields("ClassName") = Trim$(txtClassName.Text)
            .Fields("Description") = Trim$(txtDescp)
            .Update
        End With
        Call frmSelectClass.reload_rec
        rsStudent.Close
        Set rsStudent = Nothing
    End If
    '-- Prompt user.
    If g_AddClassState = 1 Then
        Dim lReply As Long
        lReply = MsgBox("Do you want to add another new class record?", vbQuestion + vbYesNo, "Add New Record")
        If lReply = vbYes Then
            Call ClearAll(Me)
            g_AddClassState = 1
            Call init
            txtClassName.SetFocus
        Else
            Unload Me
        End If
    Else
        'Call PointToRecord(rsStudent, "ID", False, g_lOldID, 0)
        Call Unload(Me)
    End If

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmFee", "cmdSave_Click", True)
    End If
    
End Sub

Private Sub Form_Load()
    Call StartBusy

    Icon = frmMain.Icon
    Call CenterForm(Me)
    
    Call initToolbar(cmdHeader, frmMain.imlTrans)
    Call initToolbar(cmdSave, frmMain.imlTrans, , g_eIcon.Save)
    Call initToolbar(cmdCancel, frmMain.imlTrans, , g_eIcon.Cancel)
    Call initToolbar(cmdHelp, frmMain.imlTrans, , g_eIcon.Help)
    Call initFrame(fraBackGround)
    
    Call init
    
    Call EndBusy
End Sub
Private Sub init()

On Error GoTo errHandler
If g_AddClassState = 1 Then
    Set rsStudent = New ADODB.Recordset
    rsStudent.Open "Select * FROM Class", g_cn, adOpenDynamic, adLockOptimistic
    rsStudent.AddNew
    'MsgBox "Class ID: " & (rsStudent!ClassID)
    txtClassID.Text = rsStudent!ClassID
    
End If
errHandler:
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
    
End Sub
