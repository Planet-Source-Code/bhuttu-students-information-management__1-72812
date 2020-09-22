VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectClass 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   4710
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
      Caption         =   "Select Class"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin prjStudentIMS.ucHorizontal3DLine ucHorizontal3DLine1 
         Height          =   30
         Left            =   480
         TabIndex        =   8
         Top             =   4080
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   53
      End
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
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton sel2 
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Add New Class"
         Top             =   455
         Width           =   315
      End
      Begin VB.CommandButton sel3 
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Edit Existing Class Record"
         Top             =   770
         Width           =   315
      End
      Begin VB.CommandButton sel4 
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Delete Class Record"
         Top             =   1085
         Width           =   315
      End
      Begin VB.CommandButton sel5 
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Refresh Class Records"
         Top             =   1400
         Width           =   315
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3510
         Left            =   480
         TabIndex        =   7
         Top             =   480
         Width           =   6540
         _ExtentX        =   11536
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Sl #"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Class ID #"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Class Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Description"
            Object.Width           =   4410
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSelectClass"
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
    With frmMain
        'sel1.Picture = .i16x16.ListImages(9).Picture
        sel2.Picture = .i16x16.ListImages(10).Picture
        sel3.Picture = .i16x16.ListImages(11).Picture
        sel4.Picture = .i16x16.ListImages(12).Picture
        sel5.Picture = .i16x16.ListImages(13).Picture
        
        Set ListView1.SmallIcons = .i16x16
        Set ListView1.Icons = .i16x16
    End With
    Call reload_rec
End Sub

Public Sub reload_rec()
    ListView1.ListItems.Clear
    Set rsTemp = New ADODB.Recordset
    rsTemp.CursorLocation = adUseClient
    rsTemp.Open "SELECT * FROM Class ORDER BY ClassID ASC", g_cn, adOpenStatic, adLockOptimistic
    rsTemp.Filter = ""
    rsTemp.Requery
    FillListView ListView1, rsTemp, 4, 2, True, True
    rsTemp.Close
    Set rsTemp = Nothing
End Sub
Private Sub selectCurList()
    If ListView1.ListItems.Count < 1 Then MsgBox "There is currently no record to select. Please add some record first.", vbExclamation, "No Records...": Exit Sub
    On Error Resume Next
    frmPatient.txtClass.Text = ListView1.SelectedItem.ListSubItems(2)
    frmPatient.txtClassID.Text = ListView1.SelectedItem.ListSubItems(1)
    Unload Me
End Sub

Private Sub ListView1_DblClick()
    Call selectCurList
End Sub

Private Sub sel2_Click()
    If user_level <> 1 Then MsgBox "You don't have the sufficient privilege to add any new class record." & vbCrLf & _
                                   "Only an administrative user can perform such a task.", vbExclamation, "Warning...": Exit Sub
    frmClass.g_AddClassState = 1
    Load frmClass
    frmClass.Caption = frmMain.Caption & " [Add New Class Record]"
    frmClass.Show vbModal
End Sub

Private Sub sel3_Click()
    If user_level <> 1 Then MsgBox "You don't have the sufficient privilege to edit any class record." & vbCrLf & _
                                   "Only an administrative user can perform such a task.", vbExclamation, "Warning...": Exit Sub
    frmClass.g_AddClassState = 0
    Load frmClass
    frmClass.txtClassID.Text = ListView1.SelectedItem.ListSubItems(1)
    frmClass.txtClassName.Text = ListView1.SelectedItem.ListSubItems(2)
    frmClass.txtDescp.Text = ListView1.SelectedItem.ListSubItems(3)
    frmClass.Caption = frmMain.Caption & " [Edit Class Record]"
    frmClass.Show vbModal
End Sub

Private Sub sel4_Click()
    
    Dim class_id As Long
    If user_level <> 1 Then MsgBox "You don't have the sufficient privilege to delete any class record." & vbCrLf & _
                                   "Only an administrative user can perform such a task.", vbExclamation, "Warning...": Exit Sub
    class_id = CLng(ListView1.SelectedItem.ListSubItems(1))
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open "SELECT * FROM Students WHERE ClassID=" & CLng(class_id) & "", g_cn, adOpenKeyset, adLockOptimistic
    If rsTemp.EOF Or rsTemp.BOF Then
        If rsTemp.State = adStateOpen Then rsTemp.Close
        rsTemp.Open "DELETE * FROM Class WHERE ClassID=" & CLng(class_id) & "", g_cn, adOpenKeyset, adLockOptimistic
        Call frmSelectClass.reload_rec
        Set rsTemp = Nothing
    ElseIf Not rsTemp.EOF Or Not rsTemp.BOF Then
        rsTemp.Close
        Set rsTemp = Nothing
        MsgBox "There are some student(s) belonging to the class selected for deleting." & vbCrLf & _
               "Hence, the class record cannot be deleted. To delete the class record" & vbCrLf & _
               "first delete all these student(s) records belonging to the class.", vbExclamation, "Cannot Delete Record..."
        Exit Sub
    
    End If
End Sub

Private Sub sel5_Click()
    Call reload_rec
End Sub
