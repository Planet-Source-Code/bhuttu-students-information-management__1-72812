Attribute VB_Name = "modMail"
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const SW_SHOWNORMAL = 1

'This function is to open default emailprogram or browser
Public Function modMailOpen(UrlMailto As String) As Boolean
    modMailOpen = ShellExecute(&O0, "Open", UrlMailto, vbNullString, vbNullString, 4)
End Function

Public Sub OpenURL(urlADD As String, sourceHWND As Long)
     Call ShellExecute(sourceHWND, vbNullString, urlADD, "", vbNullString, 1)
End Sub
