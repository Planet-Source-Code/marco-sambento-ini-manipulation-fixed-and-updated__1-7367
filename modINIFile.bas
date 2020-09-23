Attribute VB_Name = "modINIFile"
Attribute VB_Description = "INI manipulation"
'Some of these functions return the requested strings, or nonzero values,
'if success was accomplished, so you can easily set return parameters
'that allow managing the results and errors.
'
'   For example, WritePrivateProfileString returns a non zero value if it succedded
'
'On function:
'
'Public Function DeleteValue(ByVal Section As String, ByVal Key As String) As Boolean
'    WritePrivateProfileString Section, Key, "", INIFile
'End Function
'
'           we could change to:
'
'
'Public Function DeleteValue(ByVal Section As String, ByVal Key As String) As Boolean
'   Dim retval As Integer
'   retval = WritePrivateProfileString(Section, Key, "", INIFile)
'   If retval <> 0 Then
'        DeleteValue = True
'    Else
'        DeleteValue = False
'    End If
'End Function
'
'    Then, When Calling the function, like this:
'
'        Success = DeleteValue(Section, Key)
'       If Success Then '...proceed...' Else 'Msgbox error... or goto err_handler...'
'
'However, due to structure of this project, it isn't probable that any of this
'functions will fail, except when editing in notepad and not refreshing in program...

Option Explicit

Dim INIFile As String

Public Declare Function GetPrivateProfileString Lib "kernel32" _
Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, _
ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileSection Lib "kernel32" _
Alias "GetPrivateProfileSectionA" (ByVal lpApplicationName As String, _
ByVal lpReturnedString As String, _
ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString& Lib "kernel32" Alias _
"WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal Keydefault$, _
ByVal Filename$)

Public Function GetVal(ByVal Section As String, ByVal Key As String) As String
Attribute GetVal.VB_Description = "Returns the value of a key in an INI file"
    
    Dim retval As String
    Dim Result As Integer
    
    retval = String$(255, 0)
    Result = GetPrivateProfileString(Section, Key, "", retval, Len(retval), INIFile)

    If Result = 0 Then
        GetVal = vbNullString
    Else
        GetVal = Left(retval, InStr(retval, vbNullChar) - 1)
    End If
    
End Function

Public Function GetSection(ByVal Section As String) As String
Attribute GetSection.VB_Description = "Returns the 'key=value' strings from a section"

    Dim retval As String
    Dim Result As Integer
    
    retval = String$(2048, 0)
    Result = GetPrivateProfileSection(Section, retval, Len(retval), INIFile)

    If Result = 0 Then 'Failed To Retreive Section
        GetSection = vbNullString
    Else
        GetSection = Left(retval, InStr(retval, vbNullChar & vbNullChar))
    End If

End Function

Public Function Create(ByVal Section As String, ByVal Key As String, ByVal Value As String) As Boolean
Attribute Create.VB_Description = "Add To INI\r\nCreates a section or a key / value inside a section / key or creates a new section/key/value if nothing specified"
    WritePrivateProfileString Section, Key, Value, INIFile
End Function

Public Function DeleteSection(ByVal Section As String) As Boolean
Attribute DeleteSection.VB_Description = "Delete a Section\r\nDeletes a Section and all of its keys and values"
    WritePrivateProfileString Section, vbNullString, vbNullString, INIFile
End Function

Public Function DeleteKey(ByVal Section As String, ByVal Key As String) As Boolean
Attribute DeleteKey.VB_Description = "Delete a Key\r\nDeletes a key and its values"
    WritePrivateProfileString Section, Key, vbNullString, INIFile
End Function

Public Function DeleteValue(ByVal Section As String, ByVal Key As String) As Boolean
Attribute DeleteValue.VB_Description = "Delete a Value\r\nDeletes a value from a given section and key"
    WritePrivateProfileString Section, Key, "", INIFile
End Function

Public Property Let INIPath(ByVal newINI As String)
Attribute INIPath.VB_Description = "Returns/Set the INI file Path"
    INIFile = newINI
If INIFile = vbNullString Then frmMain.FileOpen = False Else frmMain.FileOpen = True
End Property

Public Property Get INIPath() As String
    INIPath = INIFile
End Property
