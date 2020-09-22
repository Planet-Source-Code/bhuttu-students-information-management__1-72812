VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSelectBookJr 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8415
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjStudentIMS.ucGradContainer fraBackGround 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8281
      Caption         =   "Select Book/Reference Title"
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
         Left            =   7080
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
         Left            =   5760
         TabIndex        =   1
         Top             =   4200
         Width           =   1215
      End
      Begin prjStudentIMS.ucHorizontal3DLine ucHorizontal3DLine1 
         Height          =   30
         Left            =   240
         TabIndex        =   3
         Top             =   4080
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   53
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3510
         Left            =   195
         TabIndex        =   4
         Top             =   480
         Width           =   7980
         _ExtentX        =   14076
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Book/Ref ID"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Book/Ref Title"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Authors(s)"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Copies Aval."
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Publishers"
            Object.Width           =   4410
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSelectBookJr"
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
    g_sSQL = "SELECT * FROM BookJournel"
    rsTemp.Open g_sSQL, g_cn, adOpenDynamic, adLockOptimistic
      
    ListView1.ListItems.Clear
    If Not rsTemp.EOF Or Not rsTemp.BOF Then
        'On Error Resume Next
        'Screen.MousePointer = vbHourglass
        Do While Not rsTemp.EOF
            On Error Resume Next
            Set itemToAdd = ListView1.ListItems.Add(, , rsTemp!BookJrID)
            itemToAdd.SubItems(1) = rsTemp!BookJrTitle
            itemToAdd.SubItems(2) = rsTemp!AuthorName
            itemToAdd.SubItems(3) = (rsTemp!NOC - rsTemp!CopiesIssued)
            itemToAdd.SubItems(4) = rsTemp!Publisher
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
    On Error Resume Next
    If ListView1.ListItems.Count < 1 Then MsgBox "There is currently no record to select. Please add some record first.", vbExclamation, "No Records...": Exit Sub
    If Val(ListView1.SelectedItem.ListSubItems(3)) = 0 Then
        MsgBox "Currently all the copies of the selected title has been issued to the students." & vbCrLf & _
               "Hence the selected title cannot be issued till one or more copie(s) of the same" & vbCrLf & _
               "is/are returned back by the student(s).", vbExclamation, "Warning..."
        Exit Sub
    End If
    frmBookJrTrans.txtTitle.Text = ListView1.SelectedItem.ListSubItems(1)
    frmBookJrTrans.txtAuthors.Text = ListView1.SelectedItem.ListSubItems(2)
    frmBookJrTrans.txtBookID.Text = ListView1.SelectedItem.Text
    Unload Me
End Sub

Private Sub ListView1_DblClick()
    Call selectCurList
End Sub

