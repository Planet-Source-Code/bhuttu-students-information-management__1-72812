VERSION 5.00
Begin VB.Form frmSearch_Staff 
   BorderStyle     =   0  'None
   Caption         =   "Fee"
   ClientHeight    =   2775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
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
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjStudentIMS.ucGradContainer fraBackGround 
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   4895
      Caption         =   "Search Staff Record"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox txtText 
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Top             =   1680
         Width           =   3255
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   4800
         TabIndex        =   1
         Top             =   2280
         Width           =   1335
      End
      Begin prjStudentIMS.ucHorizontal3DLine ucHorizontal3DLine1 
         Height          =   30
         Left            =   600
         TabIndex        =   2
         Top             =   2160
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   53
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter the staff name :"
         Height          =   195
         Left            =   600
         TabIndex        =   5
         Top             =   1680
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmFee_ae.frx":0000
         Height          =   855
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   5535
      End
   End
End
Attribute VB_Name = "frmSearch_Staff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim itemToAdd
Private Sub cmdCancel_Click()
    Call Unload(Me)
End Sub

Private Sub Form_Activate()
    txtText.SetFocus
End Sub

Private Sub Form_Load()
    
    Call CenterForm(Me)
    Icon = frmMain.Icon
    Call initFrame(fraBackGround)

End Sub

Private Sub txtText_Change()
    Call LoadLV
End Sub
Private Sub LoadLV()

    Set rsTemp = New ADODB.Recordset
    '-- Students
    g_sSQL = "SELECT * FROM Staff WHERE StaffName LIKE '%" & CStr(Trim(txtText.Text)) & "%'"
    rsTemp.Open g_sSQL, g_cn, adOpenDynamic, adLockOptimistic
      
    frmMain.lvListView.ListItems.Clear
    If Not rsTemp.EOF Or Not rsTemp.BOF Then
        'On Error Resume Next
        'Screen.MousePointer = vbHourglass
        Do While Not rsTemp.EOF
            On Error Resume Next
            Set itemToAdd = frmMain.lvListView.ListItems.Add(, , rsTemp!StaffID, , 1)
            itemToAdd.SubItems(1) = rsTemp!StaffName
            itemToAdd.SubItems(2) = Format(rsTemp!DOJ, "dd-MMM-yyyy")
            itemToAdd.SubItems(3) = rsTemp!Designation
            itemToAdd.SubItems(4) = rsTemp!Phone_No
            itemToAdd.SubItems(5) = rsTemp!MPhone_No
            itemToAdd.SubItems(6) = rsTemp!EMail
            itemToAdd.SubItems(7) = rsTemp!Qualification
            itemToAdd.SubItems(8) = rsTemp!Specialisation
            itemToAdd.SubItems(9) = rsTemp!Address
            itemToAdd.SubItems(10) = rsTemp!Remarks
            itemToAdd.SubItems(11) = rsTemp!Status
            rsTemp.MoveNext
        Loop
        DoEvents
        Call frmMain.ShowRecordInfo(rsTemp, frmMain.cmdFooter)
        rsTemp.Close
        Set rsTemp = Nothing
    Else
        Call frmMain.ShowRecordInfo(rsTemp, frmMain.cmdFooter)
        If rsTemp.State = adStateOpen Then rsTemp.Close
        Set rsTemp = Nothing
    End If
    
End Sub

