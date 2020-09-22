VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectStud 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjStudentIMS.ucGradContainer fraBackGround 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   8281
      Caption         =   "Select Student"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton Command1 
         Caption         =   "Select"
         Default         =   -1  'True
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
         Left            =   5760
         TabIndex        =   2
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
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
         TabIndex        =   1
         Top             =   4200
         Width           =   1215
      End
      Begin prjStudentIMS.ucHorizontal3DLine ucHorizontal3DLine1 
         Height          =   30
         Left            =   240
         TabIndex        =   3
         Top             =   4080
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   53
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3510
         Left            =   200
         TabIndex        =   4
         Top             =   480
         Width           =   6780
         _ExtentX        =   11959
         _ExtentY        =   6191
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Student ID"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Student Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Class"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Roll No"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Section"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Academic YR"
            Object.Width           =   3528
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSelectStud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call selectCurList
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CenterForm(Me)
    Icon = frmMain.Icon
    Call initFrame(fraBackGround)
    Call reload_rec
End Sub

Public Sub reload_rec()
    '-- Recordsets.
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    
    '-- Patient.
    g_sSQL = "SELECT * FROM qryStudents WHERE AcadYR_Duration='" & CStr(cur_acadYear) & "'"
    rsTemp.Open g_sSQL, g_cn, adOpenDynamic, adLockOptimistic
      
    ListView1.ListItems.Clear
    If Not rsTemp.EOF Or Not rsTemp.BOF Then
        'On Error Resume Next
        'Screen.MousePointer = vbHourglass
        Do While Not rsTemp.EOF
            On Error Resume Next
            Set itemToAdd = ListView1.ListItems.Add(, , rsTemp!StudentID)
            itemToAdd.SubItems(1) = rsTemp!StudName
            itemToAdd.SubItems(2) = rsTemp!ClassName
            itemToAdd.SubItems(3) = rsTemp!RollNo
            itemToAdd.SubItems(4) = rsTemp!Section
            itemToAdd.SubItems(5) = rsTemp!AcadYR_Duration
            rsTemp.MoveNext
        Loop
        DoEvents
        rsTemp.Close
        Set rsTemp = Nothing
    Else
        If rsTemp.State = adStateOpen Then rsTemp.Close
        Set rsTemp = Nothing
    End If
End Sub
Private Sub selectCurList()
    If ListView1.ListItems.Count < 1 Then MsgBox "There is currently no record to select. Please add some record first.", vbExclamation, "No Records...": Exit Sub
    On Error Resume Next
    If reqdFor_Assignmt = True Then
        frmAssignmt.txtAT.Text = ListView1.SelectedItem.ListSubItems(1)
        frmAssignmt.txtStudID.Text = ListView1.SelectedItem.Text
    ElseIf reqdFor_IntAssmt = True Then
        frmIntAssmt.txtStudent.Text = ListView1.SelectedItem.ListSubItems(1)
        frmIntAssmt.txtStudID.Text = ListView1.SelectedItem.Text
        frmIntAssmt.txtClass.Text = ListView1.SelectedItem.ListSubItems(2)
        frmIntAssmt.txtSection.Text = ListView1.SelectedItem.ListSubItems(4)
        frmIntAssmt.txtRoll.Text = ListView1.SelectedItem.ListSubItems(3)
    ElseIf reqdFor_BJrTrans = True Then
        frmBookJrTrans.txtStudent.Text = ListView1.SelectedItem.ListSubItems(1)
        frmBookJrTrans.txtStudID.Text = ListView1.SelectedItem.Text
        frmBookJrTrans.txtClass.Text = ListView1.SelectedItem.ListSubItems(2)
        frmBookJrTrans.txtSection.Text = ListView1.SelectedItem.ListSubItems(4)
        frmBookJrTrans.txtRoll.Text = ListView1.SelectedItem.ListSubItems(3)
    End If
    Unload Me
End Sub

Private Sub ListView1_DblClick()
    Call selectCurList
End Sub
