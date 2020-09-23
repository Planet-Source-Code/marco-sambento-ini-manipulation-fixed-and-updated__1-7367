Attribute VB_Name = "modMain"
Option Explicit

Public Const email = "Marco.Sambento@netc.pt?subject="
Public Const SW_SHOWNORMAL = 1
Dim WordPad As String
Global WinDir As String

Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Sub Main()
    GetWinDir
    frmMain.Show
End Sub

Function GetWinDir()

Dim WinDirectory As String  ' receives path of Windows directory
Dim slength As Long  ' receives length of the string returned

WinDirectory = Space(255)  ' initialize buffer to receive the string
slength = GetWindowsDirectory(WinDirectory, 255)  ' read the path of the Windows directory
WinDir = Left(WinDirectory, slength)  ' extract the returned string from the buffer

End Function

Function WordPadPath()
    Dim Button As VbMsgBoxResult
    Dim retval As String
    Dim Result As Integer
    
    retval = String$(255, 0)
    Result = GetPrivateProfileString("programs", "wordpad.exe", "Not Found", retval, Len(retval), WinDir & "\win.ini")
    
    If Result <> 0 Then
        WordPad = Left(retval, InStr(retval, vbNullChar) - 1)
    End If
End Function

Public Function OpenWordPad()
    Dim Button As VbMsgBoxResult
    Dim Result As Integer
    
If WordPad = vbNullString Then Call WordPadPath
If WordPad = "Not Found" Then
    Button = MsgBox("WordPad not Found!" & vbLf & "Try opening in notepad?", vbCritical + vbYesNo)
    If Button = vbYes Then WordPad = "notepad" Else Exit Function
End If

Result = ShellExecute(0&, "open", WordPad, """" & INIPath & """", vbNullString, vbMaximizedFocus) '
'here you can check if it was successfully executed

End Function

Public Sub SendEmail()
Dim Success As Long
Success = ShellExecute(0&, vbNullString, "mailto:" & email & App.Title, vbNullString, "C:\", SW_SHOWNORMAL)
End Sub
