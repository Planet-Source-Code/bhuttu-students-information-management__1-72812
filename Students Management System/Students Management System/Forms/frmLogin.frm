VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin prjStudentIMS.ucGradContainer fraBackGround 
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4260
      CaptionColor    =   -2147483630
      Caption         =   "Login Authentication"
      CaptionAlignment=   2
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.CommandButton cmdHelp 
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Click to get login help ...."
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "Login"
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
         Left            =   3000
         TabIndex        =   6
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Exit"
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
         Left            =   4200
         TabIndex        =   5
         Top             =   1920
         Width           =   1095
      End
      Begin prjStudentIMS.ucHorizontal3DLine ucHorizontal3DLine1 
         Height          =   30
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   53
      End
      Begin VB.TextBox txtPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   2520
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1320
         Width           =   2775
      End
      Begin VB.TextBox txtUserID 
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
         Left            =   2520
         TabIndex        =   2
         Top             =   960
         Width           =   2775
      End
      Begin VB.TextBox txtDate 
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
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   600
         Width           =   2775
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   240
         Picture         =   "frmLogin.frx":0000
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter password :"
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
         Left            =   1080
         TabIndex        =   9
         Top             =   1320
         Width           =   1230
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter username :"
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
         Left            =   1080
         TabIndex        =   8
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "System date :"
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
         Left            =   1080
         TabIndex        =   7
         Top             =   600
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Dim cn  As ADODB.Connection
Dim rs  As ADODB.Recordset
    
'-- Local Variables
Private m_lCtr      As Long
Private m_lCanceled As Long
Dim m_lMod          As Long     '-- Load/Free Library.
'

Private Sub cmdCancel_Click()
    m_lCanceled = 1
    Call Unload(Me)
End Sub

Private Sub cmdHelp_Click()
    MsgBox "Use any one the following username and password combination for logging in: " & vbCrLf & vbCrLf & _
           "Username: Administrator     Password: Admin" & "    [Administrative user]" & vbCrLf & _
           "Username: Demo                 Password: Demo " & "    [Non-administrative user]", vbInformation, "Login Help..."
End Sub

Private Sub cmdOk_Click()

On Error GoTo errHandler
    
    '-- System Date.
    'If Not isValidData(txtDate, "System Date", 2) Then Exit Sub
    
    '-- Original way of detecting blank username field
    'If Not isValidData(txtUserID, "User ID") Then Exit Sub
    
    '-- Password.
    If Trim$(txtUserID) = vbNullString Then
        MsgBox "Please enter your username to login.", vbExclamation, "Authentication Failed"
        txtUserID.SetFocus
        Exit Sub
    ElseIf Trim$(txtPassword) = vbNullString Then
        MsgBox "Please enter your password to login.", vbExclamation, "Authentication Failed"
        txtPassword.SetFocus
        Exit Sub
    End If
    
    'Set cn = New ADODB.Connection
    'cn.Open g_sConnectionString
    
    Set rs = New ADODB.Recordset
    g_sSQL = "SELECT * FROM tblUser WHERE UserID = '" & txtUserID & "' AND Password = '" & txtPassword & "'"
    rs.Open g_sSQL, g_cn, adOpenForwardOnly, adLockReadOnly
    
    If Not rs.EOF Or Not rs.BOF Then
        '--- Global variable
        user_level = CInt(rs!Level)
        
        With frmMain
            .cmdUser.Caption = "Log off " & Trim$(txtUserID.Text) 'g_tUser.UserID
            .cmdUser.ToolTipText = "Log off as " & Trim$(txtUserID.Text)
            '.cmdDate.Caption = txtDate
            .Show
        End With
        
        '-- Update ini file.
        'Call SaveINI(INI_NAME, "Login", "UserID", txtUserID)
        
        rs.Close
        'cn.Close
        Set rs = Nothing
        'Set cn = Nothing
        Call Unload(Me)
        
    Else
        
        '-- Increment login counter.
        m_lCtr = m_lCtr + 1
        MsgBox "Please check your username and password properly. Remember that both are case sensitive." & vbCr & vbCr & "Failed Login Attempt(s): " & m_lCtr & vbCr & "You have " & CStr(CInt(3 - m_lCtr)) & " more attempt(s) left before the program expires automatically.", vbExclamation, "Login Failed"
        If m_lCtr >= 3 Then '-- Reached maximum attempts.
            MsgBox "You have exhausted all login attempt(s). The program will now exit automatically.", vbExclamation, "Login Attempts Exhausted"
            m_lCanceled = 1
            Call Unload(Me)
        Else
            txtUserID.SetFocus
        End If
        
    End If
    
errHandler:
    Set rs = Nothing
    'Set cn = Nothing
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmLogin", "cmdOk_Click", True)
    End If
End Sub

Private Sub Form_Activate()
    txtUserID.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo errHandler

    '-- (*) KBID 309366 (http://support.microsoft.com/default.aspx?scid=kb;en-us;309366)
    m_lMod = LoadLibrary("shell32.dll")
    Call InitCommonControls
    
    Call GetGradientColor(Me.hWnd)
    
    '-- Set Login Counter and Cancel Flag.
    m_lCtr = 0
    m_lCanceled = 0

    '-- DB Connection.
    Set g_cn = New ADODB.Connection
    g_cn.Open g_sConnectionString
        
    Call CenterForm(Me)
    
    Call initFrame(fraBackGround)
    
    With frmMain
        cmdHelp.Picture = .imlTrans.ListImages(10).Picture
    End With
    '-- Default.
    txtDate = FormatDateTime$(Date, vbLongDate)
    
    'Image1.Picture = itb32x32.ListImages(1).Picture
    
    'If g_tSetting.LastLog = 1 Then
    '    txtUserID = GetINI(INI_NAME, "Login", "UserID", vbNullString)
    'End If
    
errHandler:
    If Err.Number = 364 Then '-- Object Unloaded.
        Err.Clear
        Call Unload(Me)
    ElseIf Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "frmLogin", "Form_Load", True)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call FreeLibrary(m_lMod)
    
    '-- Reset LogOff flag.
    If g_lLogOff = 1 Then g_lLogOff = 0
    
    If m_lCanceled = 1 Then
        Set rs = Nothing
        'Set cn = Nothing
        Unload Me
        End
    End If
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdOk_Click
    End If
End Sub
Private Sub txtUserID_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call cmdOk_Click
    End If
End Sub
