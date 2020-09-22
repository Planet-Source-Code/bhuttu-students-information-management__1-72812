VERSION 5.00
Object = "{BB98FE1A-DF74-4298-90F6-15DC4EC8367C}#1.0#0"; "XTab.ocx"
Begin VB.Form frmSearch_Assignms 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3900
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   ScaleHeight     =   3900
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin prjStudentIMS.ucGradContainer fraBackGround 
      Height          =   3900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   6879
      Caption         =   "Search Assignment Record"
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin prjXTab.XTab XTab1 
         Height          =   2415
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   4260
         TabCount        =   2
         TabCaption(0)   =   "     By Student Name    "
         TabContCtrlCnt(0)=   3
         Tab(0)ContCtrlCap(1)=   "txtText"
         Tab(0)ContCtrlCap(2)=   "Label1"
         Tab(0)ContCtrlCap(3)=   "Label3"
         TabCaption(1)   =   "    By Staff Name    "
         TabContCtrlCnt(1)=   3
         Tab(1)ContCtrlCap(1)=   "txtText1"
         Tab(1)ContCtrlCap(2)=   "Label4"
         Tab(1)ContCtrlCap(3)=   "Label2"
         TabStyle        =   1
         TabTheme        =   3
         ShowFocusRect   =   0   'False
         ActiveTabBackStartColor=   16316664
         ActiveTabBackEndColor=   16514555
         InActiveTabBackStartColor=   15066597
         InActiveTabBackEndColor=   15397104
         ActiveTabForeColor=   10972496
         InActiveTabForeColor=   9474192
         BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         OuterBorderColor=   9474192
         DisabledTabBackColor=   -2147483633
         DisabledTabForeColor=   -2147483627
         Begin VB.TextBox txtText1 
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
            Left            =   -72360
            TabIndex        =   7
            Top             =   1680
            Width           =   3255
         End
         Begin VB.TextBox txtText 
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
            Left            =   2640
            TabIndex        =   4
            Top             =   1680
            Width           =   3255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter the staff name :"
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
            Left            =   -74520
            TabIndex        =   9
            Top             =   1680
            Width           =   1605
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmSearch_Assignms.frx":0000
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   -74520
            TabIndex        =   8
            Top             =   540
            Width           =   5535
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   $"frmSearch_Assignms.frx":0112
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   480
            TabIndex        =   6
            Top             =   540
            Width           =   5535
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Enter the student name :"
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
            TabIndex        =   5
            Top             =   1680
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   5160
         TabIndex        =   1
         Top             =   3360
         Width           =   1335
      End
      Begin prjStudentIMS.ucHorizontal3DLine ucHorizontal3DLine1 
         Height          =   30
         Left            =   240
         TabIndex        =   3
         Top             =   3240
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   53
      End
   End
End
Attribute VB_Name = "frmSearch_Assignms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim itemToAdd
Public g_TodaysAssignms As Long
'

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
    
    If g_TodaysAssignms = 1 Then
        g_sSQL = "SELECT * FROM qryAssignments WHERE AcadYR_Duration='" & CStr(cur_acadYear) & "' AND DOS = #" & CDate(Date) & "# AND StudName LIKE '%" & CStr(Trim(txtText.Text)) & "%'"
    ElseIf g_TodaysAssignms = 0 Then
        '-- Students
        g_sSQL = "SELECT * FROM qryAssignments WHERE AcadYR_Duration='" & CStr(cur_acadYear) & "' AND StudName LIKE '%" & CStr(Trim(txtText.Text)) & "%'"
    End If
    
    rsTemp.Open g_sSQL, g_cn, adOpenDynamic, adLockOptimistic
      
    frmMain.lvListView.ListItems.Clear
    If Not rsTemp.EOF Or Not rsTemp.BOF Then
        'On Error Resume Next
        'Screen.MousePointer = vbHourglass
        Do While Not rsTemp.EOF
            On Error Resume Next
            Set itemToAdd = frmMain.lvListView.ListItems.Add(, , rsTemp!AssignID, , 1)
            itemToAdd.SubItems(1) = Format(rsTemp!DOA, "dd-MMM-yyyy")
            itemToAdd.SubItems(2) = Format(rsTemp!DOS, "dd-MMM-yyyy")
            itemToAdd.SubItems(3) = rsTemp!StudName
            itemToAdd.SubItems(4) = rsTemp!RollNo
            itemToAdd.SubItems(5) = rsTemp!ClassName
            itemToAdd.SubItems(6) = rsTemp!StaffName
            itemToAdd.SubItems(7) = rsTemp!Description
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

Private Sub txtText1_Change()
    Call LoadLV1
End Sub
Private Sub LoadLV1()

    Set rsTemp = New ADODB.Recordset
    '-- Assignments by staff
    If g_TodaysAssignms = 1 Then
        g_sSQL = "SELECT * FROM qryAssignments WHERE AcadYR_Duration='" & CStr(cur_acadYear) & "' AND DOS = #" & CDate(Date) & "# AND StaffName LIKE '%" & CStr(Trim(txtText1.Text)) & "%'"
    ElseIf g_TodaysAssignms = 0 Then
        g_sSQL = "SELECT * FROM qryAssignments WHERE AcadYR_Duration='" & CStr(cur_acadYear) & "' AND StaffName LIKE '%" & CStr(Trim(txtText1.Text)) & "%'"
    End If
    rsTemp.Open g_sSQL, g_cn, adOpenDynamic, adLockOptimistic
      
    frmMain.lvListView.ListItems.Clear
    If Not rsTemp.EOF Or Not rsTemp.BOF Then
        'On Error Resume Next
        'Screen.MousePointer = vbHourglass
        Do While Not rsTemp.EOF
            On Error Resume Next
            Set itemToAdd = frmMain.lvListView.ListItems.Add(, , rsTemp!AssignID, , 1)
            itemToAdd.SubItems(1) = Format(rsTemp!DOA, "dd-MMM-yyyy")
            itemToAdd.SubItems(2) = Format(rsTemp!DOS, "dd-MMM-yyyy")
            itemToAdd.SubItems(3) = rsTemp!StudName
            itemToAdd.SubItems(4) = rsTemp!RollNo
            itemToAdd.SubItems(5) = rsTemp!ClassName
            itemToAdd.SubItems(6) = rsTemp!StaffName
            itemToAdd.SubItems(7) = rsTemp!Description
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

Private Sub XTab1_Click()
If XTab1.ActiveTab = 0 Then
    txtText.Text = ""
    txtText.SetFocus
ElseIf XTab1.ActiveTab = 1 Then
    txtText1.Text = ""
    txtText1.SetFocus
End If
End Sub
