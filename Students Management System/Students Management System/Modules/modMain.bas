Attribute VB_Name = "modMain"

Option Explicit

'-- Global Enumerations.
Public Enum g_eMenu
    [Patient] = 1
    [Fee] = 2
    [User] = 3
    [Visitation] = 4
    [Today] = 5
    [Report] = 6
End Enum
Public Enum g_eIcon
    [Add] = 1
    [Edit] = 2
    [Delete] = 3
    [Refresh] = 4
    [Search] = 5
    [Print] = 6
    [Options] = 7
    [Save] = 8
    [Cancel] = 9
    [Help] = 10
    [Ok] = 11
    [Patient] = 1
    [Fee] = 2
    [User] = 3
    [Visitation] = 4
    [Today] = 5
    [Report] = 6
    [BOF] = 12
    [Previous] = 13
    [Next] = 14
    [EOF] = 15
End Enum

Public Enum g_eLevel
    [Personnel] = 0
    [Administrator] = 1
End Enum

'-- Local Type Declarations.
Private Type tSetting
    Splash      As Long
    LastLog     As Long
End Type

'-- Global Constants.

Public Const CPYRYT         As String = "  ProsVent Technologies © 2005 - 2006. All rights reserved. (Designed and developed by: Partha S. Paul)"

Public Const DB_NAME        As String = "MAin_Database.mdb"


'-- Global Variables.
Public cur_acadYear As String
Public rsStudent As ADODB.Recordset 'temporary recordset to manage all works
Public rsTemp As ADODB.Recordset

Public reqdFor_Assignmt As Boolean
Public reqdFor_IntAssmt As Boolean
Public reqdFor_BJrTrans As Boolean
Public user_level As Integer

Public g_sAppName           As String                   '-- Application Name.
Public g_cn                 As ADODB.Connection         '-- Connection.
Public g_rsPatient          As ADODB.Recordset          '-- Patient Recordset.
Public g_rsStaff            As ADODB.Recordset          '-- Fee Recordset.
Public g_rsUser             As ADODB.Recordset          '-- User Recordset.
Public g_sConnectionString  As String                   '-- Connection String.
Public g_sSQL               As String
Public g_tSetting           As tSetting                 '-- System Settings.
Public g_lLogOff            As Long                     '-- LogOff flag.
    
'-- Global Declarations.
Public Declare Sub InitCommonControls Lib "Comctl32.dll" ()    '-- XP UI.
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" ( _
        ByVal lpLibFileName As String) As Long
Public Declare Function FreeLibrary Lib "kernel32" ( _
        ByVal hLibModule As Long) As Long
        
'-- Local Declarations.
Private Declare Function apiShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
        (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
        ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Main()

On Error GoTo errHandler

    '-- Init global variables.
    g_sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\" & DB_NAME & ";Persist Security Info=False;"

    frmSplash.Show
errHandler:
    If Err.Number = 364 Then '-- Object Unloaded.
        Err.Clear
    ElseIf Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "modMain", "Main", True)
    End If
End Sub


Public Sub CenterForm(ByRef Frm As Form)
    Frm.Left = (Screen.Width - Frm.Width) / 2
    Frm.Top = (Screen.Height - Frm.Height) / 2
End Sub


Public Sub initFrame(fraBackGround As ucGradContainer, _
                    Optional imlIcons As ImageList, _
                    Optional lImageIdx As Long)

On Error GoTo errHandler
    
    With fraBackGround
        .BorderColor = glColorBorder
        .BackColor1 = vbButtonFace
        .BackColor2 = vbButtonFace
        .HeaderColor1 = glColorHeaderColorTwo
        .HeaderColor2 = glColorHeaderColorOne
        .CaptionAlignment = Center
        If lImageIdx <> 0 Then
            Set .Icon = imlIcons.ListImages(lImageIdx).Picture
        End If
    End With

errHandler:
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub


'//---------------------------------------------------------------------------------------
Public Sub PromptError(Optional ByVal ErrNumber As String = vbNullString, _
                        Optional ByVal ErrDescription As String = vbNullString, _
                        Optional ByVal ErrSource As String = vbNullString, _
                        Optional ByVal ModuleName As String = vbNullString, _
                        Optional ByVal ProcName As String = vbNullString, _
                        Optional ByVal DisplayPrompt As Boolean)

    Dim tString As String
    Dim fnum As Integer

    '-- Show Error Message.
    If Err.Number <> 0 Then
        If DisplayPrompt Then
            tString = "Error occured: "
            If Trim$(ErrNumber) <> "" Then tString = tString & ErrNumber & vbNewLine Else tString = tString & vbNewLine
            If Trim$(ErrDescription) <> "" Then tString = tString & "Description: " & ErrDescription & vbNewLine
            If Trim$(ErrSource) <> "" Then tString = tString & "Source: " & ErrSource & vbNewLine
            If Trim$(ModuleName) <> "" Then tString = tString & "Module: " & ModuleName & vbNewLine
            If Trim$(ProcName) <> "" Then tString = tString & "Function: " & ProcName
            MsgBox tString, vbCritical, App.Title
        End If
        '-- Write error log.
        fnum = FreeFile
        Open App.Path & "\ErrorLog.txt" For Append As #fnum
        Write #fnum, Date, ErrNumber, ErrDescription, ErrSource, ModuleName, ProcName, Environ("username"), Environ("computername")
        Close #fnum
    End If

End Sub

'//---------------------------------------------------------------------------------------

Public Sub initToolbar(ByRef cmdLavolpe As lvButtons_H, _
                        imlIcons As ImageList, _
                        Optional lFooter As Long, _
                        Optional lPictureIdx As Long, _
                        Optional lPictureSize As Long = lv_16x16)

    With cmdLavolpe
        If lFooter = 0 Then '-- Header.
            .BackColor = glColorTwoNormal
            .GradientColor = glColorOneNormal
            .HoverBackColor = glColorTwoSelected
            .HoverBackColorEnd = glColorOneSelected
            .PictureSize = lPictureSize
        Else    '-- Footer.
            .BackColor = glColorOneNormal
            .GradientColor = glColorTwoNormal
            .HoverBackColor = glColorOneSelected
            .HoverBackColorEnd = glColorTwoSelected
            .PictureSize = lPictureSize
        End If

        '-- Button picture.
        If lPictureIdx > 0 Then
            Set .Picture = imlIcons.ListImages(lPictureIdx).Picture
        End If

        .GradientMode = lv_Bottom2Top
        .ButtonStyle = lv_hover
        .PictureAlign = lv_LeftOfCaption
    End With

End Sub

'-Notes     : Lookup
'//---------------------------------------------------------------------------------------
'
Public Sub initButton(ByRef cmdLavolpe As lvButtons_H, _
                        Optional imlIcons As ImageList, _
                        Optional lPictureIdx As Long, _
                        Optional lPictureSize As Long = lv_16x16)
                        
    Set cmdLavolpe.Picture = imlIcons.ListImages(lPictureIdx).Picture
    
    cmdLavolpe.ButtonStyle = lv_hover
    cmdLavolpe.PictureSize = lPictureSize
    cmdLavolpe.ShowFocusRect = False
    'Set cmdLavolpe.MouseIcon = LoadResPicture(101, 2)
    
End Sub

'-------------------------------------------------------------------------
Public Sub StartBusy()
    Screen.MousePointer = vbHourglass                   '-- Change screen mouse cursor.
    'frmMain.ucWait.PlayWait                             '-- Play wait control.
    frmMain.ucStatus.PanelText(1) = "Processing..."       '-- Change main form caption.
End Sub

'-----------------------------------------------------------------------------------
Public Sub EndBusy()
    Screen.MousePointer = vbDefault                     '-- Change screen mouse cursor.
    'frmMain.ucWait.StopWait                             '-- Play wait control.
    frmMain.ucStatus.PanelText(1) = CPYRYT               '-- Change main form caption.
End Sub


Public Sub ClearAll(Frm As Form)
    Dim ctl As Control
    For Each ctl In Frm
        If TypeOf ctl Is TextBox Then
            ctl.Text = vbNullString
        ElseIf TypeOf ctl Is ComboBox Then
            ctl.ListIndex = -1
        End If
    Next
End Sub


Public Sub Numeric(txt As TextBox, KeyAscii As Integer)
    Select Case KeyAscii
        Case "48" To "57", vbKeyBack                        '-- Numbers 0-9, Backspace
        Case "46"                                           '-- Period(.) ,45 -> Negative Sign
            If Len(txt) = txt.SelLength Then Exit Sub       '-- Check if whole text is highlighted.
            If InStr(1, txt, ".") > 0 Then KeyAscii = 0     '-- Lock Decimal.
        Case "13"                                           '-- Enter Key.
            SendKeys "{tab}"
        Case Else
            KeyAscii = 0
    End Select
End Sub

Public Sub Navigate(ByVal Index As Integer, _
                    ByRef rs As ADODB.Recordset)

On Error GoTo errHandler

    If Not (rs.BOF And rs.EOF) Then
        Select Case Index
        Case 0 '-- First
            rs.MoveFirst
        Case 1 '-- Previous
            rs.MovePrevious
            If rs.BOF Then
                rs.MoveFirst
            End If
        Case 2 '-- Next
            If rs.EOF Then
                rs.MoveLast
            Else
                rs.MoveNext
                If rs.EOF Then
                    rs.MoveLast
                End If
            End If
        Case 3 '-- Last
            rs.MoveLast
        End Select
    End If

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "modGlobal", "Navigate", True)
    End If
End Sub


Public Function isValidData(ctl As Control, sDesc As String, Optional lType As Long) As Boolean
    
On Error GoTo errHandler

    If TypeOf ctl Is ComboBox Then '-- Combobox.
        '-- Check for empty values
        If ctl.ListIndex = -1 Then
            MsgBox "Please select a valid " & sDesc & ".", vbExclamation + vbOKOnly, "Incomplete Data..."
            ctl.SetFocus
            Exit Function
        End If
    Else '-- Textbox.
        '-- Check for empty values.
        If Trim$(ctl) = vbNullString Then
            MsgBox "Please enter " & sDesc & ".", vbExclamation + vbOKOnly, "Incomplete Data..."
            ctl.SetFocus
            Exit Function
        End If
    End If
    
    If lType = 1 Then '-- Number.
        If Not IsNumeric(ctl) Then
            MsgBox sDesc & " is not a valid number.", vbExclamation + vbOKOnly
            ctl.SetFocus
            Exit Function
        End If
    ElseIf lType = 2 Then '-- Date.
        If Not IsDate(ctl) Then
            MsgBox sDesc & " is not a valid date.", vbExclamation + vbOKOnly
            ctl.SetFocus
            Exit Function
        End If
    End If
    
    isValidData = True

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "modMain", "isValidData", True)
    End If

End Function


Public Function OpenThisFile(ByVal stFile As String, _
                        ByVal lShowHow As Long, _
                        ByVal sParams As String, _
                        ByRef lhWnd As Long) As Variant

    Dim lRet As Long, stRet As String, ErrID As Long

On Error GoTo TryAPIcall
    lRet = -1   ' set default value -- meaning failure
    If Len(sParams) > 0 Then sParams = " " & sParams    ' if no optional parameters, then format with a space
    lRet = Shell(stFile & sParams, lShowHow)       ' attempt simple shell command
    OpenThisFile = lRet
    Exit Function

TryAPIcall:
    Err.Clear
    ' if above shell function failed, then try an association open based on the file extension
    ErrID = apiShellExecute(lhWnd, "OPEN", _
            stFile, sParams, App.Path, lShowHow)
    ' Errors will be a retruned value of <32
    If ErrID < 32& Then
        Select Case ErrID
            Case 31&:
                'Try the OpenWith dialog
                lRet = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " _
                        & stFile, 1)
            Case 0&:
                stRet = "Out of Memory/Resources. Could not Execute."
            Case 2&:
                stRet = "File not found. Could not Execute."
            Case 3&:
                stRet = "Path not found. Could not Execute."
            Case 11&:
                stRet = "Bad File Format. Could not Execute."
            Case Else:
        End Select
        If ErrID <> 31 Then
            lRet = -1 ' failure
            MsgBox stRet, vbExclamation + vbOKOnly  ' display error
        End If
        OpenThisFile = lRet
    Else
        lRet = 69
    End If
Resume Next
End Function


Public Function isExisting(cn As ADODB.Connection, sTableName As String, sCondition As String) As Boolean
    Dim rs As ADODB.Recordset
    
On Error GoTo errHandler

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM " & sTableName & " WHERE " & sCondition, cn, adOpenForwardOnly, adLockReadOnly
    If Not rs.EOF Then isExisting = True
    
errHandler:
    Set rs = Nothing
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "modDB", "isExisting", True)
    End If
End Function


Public Sub PointToRecord(ByRef rs As ADODB.Recordset, ByVal sField As String, ByVal isString As Boolean, ByVal sStr As String, ByVal sNum As Long)
    Dim lOldPosition As Long
    Dim sqlParam As String
    
On Error GoTo errHandler

    With rs
        lOldPosition = .AbsolutePosition
        rs.Filter = adFilterNone
        rs.Requery
        .MoveFirst
        '/Check if string or number.
        If isString Then
            sqlParam = sField & " = '" & sStr & "'"
        Else
            sqlParam = sField & " = " & sNum
        End If
        .Find sqlParam
        If .EOF Then .AbsolutePosition = lOldPosition
    End With

errHandler:
    If Err.Number <> 0 Then
        Call PromptError(Err.Number, Err.Description, Err.Source, "modDB", "PointToRecord", True)
    End If
End Sub


'Procedure used to fill list view
Public Sub FillListView(ByRef sListView As ListView, ByRef sRecordSource As Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional srcHiddenField As String)
    Dim X As Variant
    Dim i As Byte
    On Error Resume Next
    sListView.ListItems.Clear
    If sRecordSource.RecordCount < 1 Then Exit Sub
    sRecordSource.MoveFirst
    Do While Not sRecordSource.EOF
        If with_num = True Then
            Set X = sListView.ListItems.Add(, , sRecordSource.AbsolutePosition, sNumIco, sNumIco)
        Else
            Set X = sListView.ListItems.Add(, , "" & sRecordSource.Fields(0), sNumIco, sNumIco)
        End If
            If srcHiddenField <> "" Then X.Tag = sRecordSource.Fields(srcHiddenField)
            For i = 1 To sNumOfFields - 1
                If show_first_rec = True Then
                    If with_num = True Then
                        If sRecordSource.Fields(CInt(i) - 1).Type = adDouble Then
                            X.SubItems(i) = FormatRS(sRecordSource.Fields(CInt(i) - 1))
                        Else
                            X.SubItems(i) = "" & FormatRS(sRecordSource.Fields(CInt(i) - 1))
                        End If
                    Else
                        If sRecordSource.Fields(CInt(i)).Type = adDouble Then
                            X.SubItems(i) = FormatRS(sRecordSource.Fields(CInt(i)))
                        Else
                            X.SubItems(i) = "" & FormatRS(sRecordSource.Fields(CInt(i)))
                        End If
                    End If
                Else
                    X.SubItems(i) = "" & FormatRS(sRecordSource.Fields(CInt(i) + 1))
                End If
            Next i
        sRecordSource.MoveNext
    Loop
    i = 0
    Set X = Nothing
End Sub
'Function used to format recordset
Public Function FormatRS(ByVal srcField As Field, Optional AllowNewLine As Boolean) As String
    Dim strRet As String
    
    With srcField
        If AllowNewLine = True Then
            strRet = srcField
        Else
            strRet = Replace(srcField, vbCrLf, " ", , , vbTextCompare)
        End If
        
        If srcField.Type = adCurrency Or srcField.Type = adDouble Then
            strRet = Format$(srcField, "#,##0.00")
        ElseIf srcField.Type = adDate Then
            strRet = Format$(srcField, "MMM-dd-yyyy")
        Else
            strRet = srcField
        End If
    End With
    
    FormatRS = strRet
    
    strRet = vbNullString
End Function
