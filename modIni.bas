Attribute VB_Name = "modIni"
Option Explicit

'ripped from a class i wrote, and scaled down

'********************************************
'    Name:    INI File Functions
'  Author:    Lewis Miller (aka Deth)
' Purpose:    Makes reading and writing to ini files a breeze
'********************************************

'API for writing to ini files
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'API for reading from ini files
Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

Private INI_BUFFER As String

Private mvarFilePath As String
Private mvarBufferSize As Long

Public Property Let BufferSize(ByVal vData As Long)

    mvarBufferSize = vData
    INI_BUFFER = Space$(mvarBufferSize)

End Property
Public Property Get BufferSize() As Long

    BufferSize = mvarBufferSize

End Property
'this is some ending comment _

Public Sub Ini_Initialize(Optional ByVal BuffSize As Long = 255, Optional ByVal IniPath As String)

    'This supplys a default file path for a normal program.
    'To change it, set the filepath property to whatever you wish.
    If Len(IniPath) = 0 Then
        mvarFilePath = LCase$(App.Path) & "\" & LCase$(App.EXEName) & ".ini"
    Else
        mvarFilePath = IniPath
    End If
    
    BufferSize = BuffSize

End Sub

'read a number from a setting in an ini file
Public Function ReadNumber(ByVal strSection As String, ByVal strKey As String, Optional ByVal lngDefault As Long, Optional ByVal strFilePath As String) As Long

    On Error GoTo NoValue
    If LenB(strFilePath) > 0 Then mvarFilePath = strFilePath
    ReadNumber = GetPrivateProfileInt(strSection, strKey, lngDefault, mvarFilePath)
NoValue:

End Function

'write a number to an ini file
Public Sub WriteNumber(ByVal strSection As String, ByVal strKey As String, ByVal strValue As Long, Optional ByVal strFilePath As String)

    If LenB(strFilePath) > 0 Then mvarFilePath = strFilePath
    Call WritePrivateProfileString(strSection, strKey, CStr(strValue), mvarFilePath)

End Sub

'read a string from a setting in an ini file
Public Function ReadString(ByVal strSection As String, ByVal strKey As String, Optional ByVal strDefault As String, Optional ByVal strFilePath As String) As String

    On Error GoTo NoValue
    If LenB(strFilePath) > 0 Then mvarFilePath = strFilePath
    ReadString = INI_BUFFER
    ReadString = Left$(ReadString, GetPrivateProfileString(strSection, strKey, strDefault, ReadString, mvarBufferSize, mvarFilePath))
NoValue:

End Function

'write a string to an ini file
Public Sub WriteString(ByVal strSection As String, ByVal strKey As String, ByVal strValue As String, Optional ByVal strFilePath As String)

    If LenB(strFilePath) > 0 Then mvarFilePath = strFilePath
    Call WritePrivateProfileString(strSection, strKey, strValue, mvarFilePath)

End Sub



