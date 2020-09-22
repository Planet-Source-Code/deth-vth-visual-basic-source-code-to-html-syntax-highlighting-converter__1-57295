Attribute VB_Name = "modFile"
Option Explicit
'some file manipulation functions, this is a pre-requisite file i include
'in most of my projects, usually i go thru an strip unused functions

'All Code Written (Or *Heavily* Modified) by Lewis Miller

'the following api calls and structures are used within this module
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetOpenFileName Lib "COMDLG32" Alias "GetOpenFileNameA" (File As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "COMDLG32" Alias "GetSaveFileNameA" (File As OPENFILENAME) As Long
Private Declare Function ChooseColor Lib "COMDLG32.DLL" Alias "ChooseColorA" (Color As TCHOOSECOLOR) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private Type TCHOOSECOLOR
    lStructSize     As Long
    hwndOwner       As Long
    hInstance       As Long
    rgbResult       As Long
    lpCustColors    As Long
    Flags           As Long
    lCustData       As Long
    lpfnHook        As Long
    lpTemplateName  As Long
End Type

Private Type OPENFILENAME
    lStructSize             As Long
    hwndOwner               As Long
    hInstance               As Long
    lpstrFilter             As String
    lpstrCusFilt(7)         As Byte
    nFilterIndex            As Long
    lpstrFile               As String
    nMaxFile                As Long
    lpstrFileTitle          As String
    nMaxFileTitle           As Long
    lpstrInitialDir         As String
    lpstrTitle              As String
    lpFlags                 As Long
    nJunk(3)                As Byte
    lpstrDefExt             As String
    lCustData(11)           As Byte
End Type

'the next 4 Show#### functions negate the need in most cases
'to include the bastard child "comdlg32.ocx" with your application.
'function Arguments are all optional and should be self explanatory...

'ShowOpen() 'shows the open file dialog window
'ShowSave() 'shows the save file dialog window
'ShowColor() 'shows the choose color window
'showfont() and showprint() dialog not yet done...

'function relies on GetOpenFileName() api call and OPENFILENAME structure
Function ShowOpen(Optional ByVal strFilter As String, Optional ByVal strDefaultPath As String, Optional ByVal strTitle As String, Optional ByVal strDefaultExt As String, Optional ByVal strInitDir As String) As String
    'coded by Lewis Miller
    Dim OpFile As OPENFILENAME

    On Error Resume Next
    With OpFile
        InitOFStruct OpFile, strFilter, strTitle, strDefaultExt, strInitDir
        If GetOpenFileName(OpFile) = 0 Then
            ShowOpen = strDefaultPath
            Exit Function
        End If
        .nMaxFile = InStr(.lpstrFile, vbNullChar)    'instead of dimensioning an uneeded variable, we just re-use one (.nmaxfile) from the structure
        If .nMaxFile > 0 Then
            .lpstrFile = Left$(.lpstrFile, .nMaxFile - 1)
            If LenB(.lpstrFile) > 0 Then
                ShowOpen = .lpstrFile
                Exit Function
            End If
        End If
        ShowOpen = strDefaultPath
    End With

    On Error GoTo 0

End Function

'function relies on GetSaveFileName() api call and OPENFILENAME structure
Function ShowSave(Optional ByVal strFilter As String, Optional ByVal strDefaultPath As String, Optional ByVal strTitle As String, Optional ByVal strDefaultExt As String, Optional ByVal strInitDir As String) As String
    'coded by Lewis Miller
    Dim OpFile As OPENFILENAME

    On Error Resume Next
    With OpFile
        InitOFStruct OpFile, strFilter, strTitle, strDefaultExt, strInitDir
        If GetSaveFileName(OpFile) = 0 Then
            ShowSave = strDefaultPath
            Exit Function
        End If
        .nMaxFile = InStr(.lpstrFile, vbNullChar)
        If .nMaxFile > 0 Then
            .lpstrFile = Left$(.lpstrFile, .nMaxFile - 1)
            If LenB(.lpstrFile) > 0 Then
                ShowSave = .lpstrFile
                Exit Function
            End If
        End If
        ShowSave = strDefaultPath
    End With

    On Error GoTo 0

End Function

'this is a helper function for the ShowOpen and ShowSave functions.
'initializes the OPENFILENAME structure.
'this function relies on the GetActiveWindow() api call and OPENFILENAME structure
Private Sub InitOFStruct(opf As OPENFILENAME, ByVal strFilter As String, ByVal strTitle As String, ByVal strDefaultExt As String, ByVal strInitDir As String)
    'coded by Lewis Miller

    If LenB(strInitDir) = 0 Then strInitDir = App.Path
    If LenB(strTitle) = 0 Then strTitle = "Select File..."
    If LenB(strFilter) = 0 Then strFilter = "All Files (*.*)|*.*"
    With opf
        .lStructSize = Len(opf)
        .hwndOwner = GetActiveWindow
        .lpstrInitialDir = strInitDir
        .lpstrDefExt = strDefaultExt
        .lpstrTitle = strTitle
        .lpstrFilter = Replace$(Replace$(strFilter, "|", vbNullChar), ":", vbNullChar) & vbNullChar & vbNullChar
        .nFilterIndex = 1
        .nMaxFile = 280
        .nMaxFileTitle = .nMaxFile
        .lpstrFile = String$(.nMaxFile, vbNullChar)
        .lpstrFileTitle = .lpstrFile
    End With

End Sub

'function relies on ChooseColor(), GetActiveWindow() api calls and TCHOOSECOLOR structure
Public Function ShowColor(Optional ByVal DefColor As Long) As Long
    'coded by Lewis Miller

    Dim UdtChooseCol     As TCHOOSECOLOR
    Dim CustomColors(15) As Long
    Dim X                As Long

    On Error Resume Next
    With UdtChooseCol
        .lStructSize = Len(UdtChooseCol)
        .hwndOwner = GetActiveWindow
        .rgbResult = DefColor
        .Flags = 129 '(Not (CC_ENABLEHOOK Or CC_ENABLETEMPLATE)) And (CC_RGBINIT Or CC_SOLIDCOLOR)
        For X = 0 To 15
            CustomColors(X) = GetSysColor(X)
        Next X
        .lpCustColors = VarPtr(CustomColors(0))
        If ChooseColor(UdtChooseCol) Then
            ShowColor = .rgbResult
        Else
            ShowColor = -1
        End If
    End With
    On Error GoTo 0

End Function


'function relies on shellexecute() and GetActiveWindow api calls
Function ExecuteFile(ByVal strFilePath As String, Optional ByVal strVerb As String = "open", Optional ByVal strWorkingDir As String, Optional ByVal strCommandLine As String, Optional ByVal lngWindowState As Long = 1) As Boolean
    'coded by Lewis Miller

    Dim lngReturnValue As Long


    If Len(strFilePath) > 1 Then
        strFilePath = Trim$(strFilePath)

        If Len(strWorkingDir) = 0 Then    'no folder path
            'make a folder path from file path
            lngReturnValue = InStrRev(strFilePath, "\")
            If lngReturnValue > 0 Then
                strWorkingDir = Left$(strFilePath, lngReturnValue - 1)
                If Len(strWorkingDir) = 2 Then    'C:
                    strWorkingDir = strWorkingDir & "\"
                End If
            Else
                strWorkingDir = CurDir$    'use default
            End If
        End If

        If Len(strVerb) = 0 Then strVerb = "open"

        'quote filepath to prevent long filename errors
        If Left$(strFilePath, 1) <> """" And Right$(strFilePath, 1) <> """" Then
            strFilePath = """" & strFilePath & """"
        End If

        lngReturnValue = ShellExecute(GetActiveWindow, strVerb, strFilePath, strCommandLine, strWorkingDir, lngWindowState)

    Else
        'fail, no filepath...
        lngReturnValue = 2
    End If

    If lngReturnValue <= 32 Then
        '        Select Case lngReturnValue
        '            Case 2, 3
        '                MsgBox "The file was not found. Please make sure that the file is in the folder.", vbCritical
        '            Case 8
        '                MsgBox "There was insufficient memory to execute the file.", vbCritical
        '            Case Else
        '                MsgBox "An unknown error occured when trying to start the program.", vbCritical
        '        End Select
    Else
        ExecuteFile = True
    End If


End Function

'error free wrapper for the filelen() function, returns 0 if file not found
Function GetFileLength(ByVal strFilePath As String) As Long
    'coded by Lewis Miller

    On Error Resume Next
    GetFileLength = FileLen(strFilePath)
    On Error GoTo 0

End Function

'delete a file
Sub FileKill(ByVal strFilePath As String)

    On Error Resume Next
    SetAttr strFilePath, vbNormal
    Kill strFilePath
    On Error GoTo 0

End Sub

'checks to see if a file exists
Function FileExist(ByVal strFilePath As String) As Boolean
    'coded by Lewis Miller

    On Error GoTo ErrHandle

    If Not ((GetAttr(strFilePath) And vbDirectory) = vbDirectory) Then
        FileExist = (Dir$(strFilePath) <> vbNullString)
    End If

ErrHandle:
    On Error GoTo 0

End Function

'checks to see if a file is readable (can also be used as a fileexist)
Function FileCanRead(ByVal strFilePath As String) As Boolean
    'coded by Lewis Miller

    Dim intFileNum As Integer
    intFileNum = FreeFile
    If Len(strFilePath) > 2 Then
        On Error Resume Next
        Open strFilePath For Input As #intFileNum
        FileCanRead = (Err = 0)
        Close #intFileNum
        On Error GoTo 0
    End If

End Function

'dumps a file's contents into a string
Function FileToString(ByVal strFilePath As String) As String
    'coded by Lewis Miller

    Dim intFileNum As Integer

    If FileCanRead(strFilePath) Then
        On Error Resume Next
        intFileNum = FreeFile
        Open strFilePath For Binary As #intFileNum
        FileToString = Space$(LOF(intFileNum))
        Get #intFileNum, , FileToString
        Close #intFileNum
        On Error GoTo 0
    End If

End Function

'saves a string to a file as is (overwrites any previous file data)
Sub FileSave(ByVal strFileBuffer As String, ByVal strFilePath As String)
    'coded by Lewis Miller

    Dim intFileNum As Integer

    If FileCanRead(strFilePath) Then FileKill strFilePath
    intFileNum = FreeFile
    Open strFilePath For Binary As #intFileNum
    Put #intFileNum, , strFileBuffer
    Close #intFileNum

End Sub

'saves a byte array to a file (overwrites any previous file data)
Sub FileByte(Bytes() As Byte, ByVal strFilePath As String)
    'coded by Lewis Miller

    Dim intFileNum As Integer

    If FileCanRead(strFilePath) Then FileKill strFilePath
    intFileNum = FreeFile
    Open strFilePath For Binary As #intFileNum
    Put #intFileNum, , Bytes
    Close #intFileNum

End Sub


'appends a string to a file, best for log files
Sub FileAppend(ByVal strFileBuffer As String, ByVal strFilePath As String)
    'coded by Lewis Miller

    Dim intFileNum As Integer
    intFileNum = FreeFile
    Open strFilePath For Append As #intFileNum
    Print #intFileNum, strFileBuffer
    Close #intFileNum

End Sub

'writes a string to a file, and appends a CR-LF to the string
'similar to file append but it overwrites any existing data
Sub FileWrite(ByVal strFileBuffer As String, ByVal FilePath As String)
    'coded by Lewis Miller

    Dim intFileNum As Integer
    intFileNum = FreeFile
    Open FilePath For Output As #intFileNum
    Print #intFileNum, strFileBuffer
    Close #intFileNum

End Sub


'adds a new file extension to a file name or path
'can also remove extension (and dot) by making strNewExtension = ""
Function FileAddExtension(ByVal strFilePath As String, ByVal strNewExtension As String) As String
    'coded by Lewis Miller

    Dim lngPathLength As Long, lngFindPlace As Long

    lngPathLength = Len(strFilePath)
    If lngPathLength > 0 Then
        lngFindPlace = InStrRev(strFilePath, "\")
        If (lngFindPlace > 0) And (lngFindPlace < lngPathLength) Then
            FileAddExtension = Left$(strFilePath, lngFindPlace)
            strFilePath = Mid$(strFilePath, lngFindPlace + 1)
        End If
        lngFindPlace = InStrRev(strFilePath, ".")
        If lngFindPlace > 0 Then
            strFilePath = Left$(strFilePath, lngFindPlace - 1)
        End If
    End If
    If Len(strNewExtension) > 0 Then
        If Left$(strNewExtension, 1) <> "." Then
            strNewExtension = "." & strNewExtension
        End If
    End If
    FileAddExtension = FileAddExtension & strFilePath & strNewExtension

End Function

'grabs folder path from "strFilePath"
'Note: this must be passed a file path, a folder path with no trailing slash
'      will be treated as a file path with no extension. Also drives with
'      trailing slashes will be returned without them, ex: getfolderpath("C:\") will = "C:"
Function GetFolderpath(ByVal strFilePath As String) As String
    'coded by Lewis Miller

    Dim lngPathLength As Long, lngSlashMark As Long

    lngPathLength = Len(strFilePath)

    If lngPathLength > 2 Then
        If Right$(strFilePath, 1) = "\" Then 'trim any trailing slashes and return
            Do While (lngPathLength > 0) And (Right$(strFilePath, 1) = "\")
                strFilePath = Left$(strFilePath, lngPathLength - 1)
                lngPathLength = Len(strFilePath)
            Loop
            GetFolderpath = strFilePath
            Exit Function
        End If
    End If
    If lngPathLength > 0 Then
        lngSlashMark = InStrRev(strFilePath, "\")
        If lngSlashMark > 0 Then
            strFilePath = Left$(strFilePath, lngSlashMark - 1)
        End If
    End If

    GetFolderpath = strFilePath

End Function

'grabs file name from file path
Function GetFileName(ByVal strFilePath As String) As String
    'coded by Lewis Miller

    Dim lngPathLen As Long, lngSlashMark As Long

    lngPathLen = Len(strFilePath)
    If lngPathLen > 0 Then
        If Right$(strFilePath, 1) = "\" Then
            Exit Function
        End If
        lngSlashMark = InStr(strFilePath, "\")
        If lngSlashMark > 0 And lngSlashMark < lngPathLen Then
            strFilePath = Mid$(strFilePath, InStrRev(strFilePath, "\") + 1)
        End If
    End If
    GetFileName = strFilePath

End Function

'checks to see if a folder exists
Function FolderExist(ByVal strFolderPath As String) As Boolean
    'coded by Lewis Miller

    On Error Resume Next
    If Len(strFolderPath) > 1 Then
        If Right$(strFolderPath, 1) = "\" Then
            strFolderPath = strFolderPath & "nul"    'microsoft will probably remove this little MSDOS trick eventually
        Else
            strFolderPath = strFolderPath & "\nul"
        End If
        FolderExist = (Dir$(strFolderPath) <> vbNullString)
    End If
    On Error GoTo 0

End Function

'error trapped wrapper to create a folder
Sub FolderCreate(ByVal strFolderPath As String)
    'coded by Lewis Miller

    On Error Resume Next
    MkDir strFolderPath
    On Error GoTo 0

End Sub

Sub MakeDeepDir(ByVal strFolderPath As String)
    'coded by Lewis Miller

    Dim strTempFolder As String, lngNextPlace As Long

    lngNextPlace = InStr(strFolderPath, "\")
    Do While lngNextPlace > 0
        strTempFolder = strTempFolder & Left$(strFolderPath, lngNextPlace)
        If Not FolderExist(strTempFolder) Then
            FolderCreate strTempFolder
        End If
        strFolderPath = Mid$(strFolderPath, lngNextPlace + 1)
        lngNextPlace = InStr(strFolderPath, "\")
    Loop

    If Len(strFolderPath) > 0 Then
        strTempFolder = strTempFolder & strFolderPath
        FolderCreate strTempFolder
    End If

End Sub


'deletes a folder including any subfolders
Sub FolderNuke(ByVal strFolderPath As String)

    On Error Resume Next
    Dim FolderName As String

    If Right$(strFolderPath, 1) = "\" Then
        strFolderPath = Left$(strFolderPath, Len(strFolderPath) - 1)
    End If
    
    FileKill strFolderPath & "\*.*"

    Do
        FolderName = Dir$(strFolderPath & "\*.*", 16)
        While FolderName = "." Or FolderName = ".."
            FolderName = Dir$
        Wend
        If Len(FolderName) = 0 Then Exit Do
        FolderNuke strFolderPath & "\" & FolderName
    Loop

    RmDir strFolderPath
    On Error GoTo 0

End Sub

'simple function that does some checking of a strFolderPath
'and removes trailing slashes
Function FormatPath(ByVal strFolderPath As String) As String
    'coded by Lewis Miller

    On Error Resume Next
    If Len(strFolderPath) > 2 Then
        Do While Right$(strFolderPath, 1) = "\"
            strFolderPath = Left$(strFolderPath, Len(strFolderPath) - 1)
        Loop
        strFolderPath = Replace$(strFolderPath, "/", "\")
    End If

    If Len(strFolderPath) > 2 Then
        FormatPath = Left$(strFolderPath, 2) & Replace$(Mid$(strFolderPath, 3), "\\", "\")
    Else
        FormatPath = strFolderPath
    End If

End Function

'formats a filesize into a KB string e.g "100 Kb"
Function FormatBytes(ByVal dblByteCount As Double) As String

    If dblByteCount >= 1024000000# Then
        FormatBytes = Format$(dblByteCount / 1024000000, "###,###,###,###,##0.00") & " Gb"
    ElseIf dblByteCount >= 1024000# Then
        FormatBytes = Format$(dblByteCount / 1024000, "###,###,##0.00") & " Mb"
    ElseIf dblByteCount >= 1024# Then
        FormatBytes = Format$(dblByteCount / 1024, "0.00") & " Kb"
    Else
        FormatBytes = CStr(dblByteCount) & " Bytes"
    End If

End Function

'calulates hours,minutes and seconds from number of milliseconds
'use this with the timeGetTime() api call (Timer wont work very well in most cases)
Function CalculateTime(ByVal sngTotalMillsec As Single) As String
    'coded by Lewis Miller

    Dim sngHour As Single, sngMinutes As Single, sngSeconds As Single

    If sngTotalMillsec >= 3600000 Then
        sngHour = sngTotalMillsec \ 3600000
        sngTotalMillsec = sngTotalMillsec - (sngHour * 3600000)
        CalculateTime = CStr(sngHour) & " Hr. "
    End If

    If sngTotalMillsec >= 60000 Then
        sngMinutes = sngTotalMillsec \ 60000
        sngTotalMillsec = sngTotalMillsec - (sngMinutes * 60000)
        CalculateTime = CalculateTime & CStr(sngMinutes) & " Min. "
    End If

    If sngTotalMillsec >= 1000 Then
        sngSeconds = sngTotalMillsec \ 1000
        CalculateTime = CalculateTime & CStr(sngSeconds) & " Sec."
    Else
        CalculateTime = CalculateTime & ".0" & Left$(CStr(sngTotalMillsec), 1) & " Sec."
    End If

End Function

'cheap way to convert a string to a picture,
'uses/depends on the FileSave() function
Function StrToPic(ByVal strPicString As String) As IPictureDisp
    'coded by Lewis Miller

    Dim FilePath As String

    If Len(strPicString) > 0 Then
        FilePath = App.Path & "\temp.tmp"
        FileSave strPicString, FilePath
        On Error Resume Next
        Set StrToPic = LoadPicture(FilePath)
        On Error GoTo 0
    End If

End Function

'makes a composite file path from a base folder path and a relative file path
'assumes correct input...
Function AbsoluteFromRelative(ByVal strBaseDir As String, ByVal strFilePath As String) As String
    'coded by Lewis Miller

    'C:\Code\My Progs\my program
    '      ..\      ..\        ..\WINNT\System32\stdole2.tlb

    If InStr(strFilePath, "..\") Then
        On Error Resume Next    'just in case
        Do While Left$(strFilePath, 3) = "..\" And Len(strBaseDir) > 2
            strBaseDir = Left$(strBaseDir, InStrRev(strBaseDir, "\") - 1)
            strFilePath = Mid$(strFilePath, 4)
        Loop
        On Error GoTo 0
    End If

    AbsoluteFromRelative = strBaseDir & "\" & strFilePath

End Function

'creates a filename that doesnt exist by adding an number to filename, ex: filename[2].ext
'this function works with a folderpath and a file name.
Function SafeFileName(ByVal strFolderPath As String, ByVal strFileName As String) As String
    'coded by Lewis Miller

      Dim strFileExt    As String
      Dim lngCount      As Long
      Dim lngPlace      As Long
      
    On Error Resume Next
        SafeFileName = strFileName
        If InStr(strFileName, ".") Then
            lngPlace = InStrRev(strFileName, ".")
            strFileExt = Mid$(strFileName, lngPlace)
            strFileName = Left$(strFileName, lngPlace - 1)
        End If
        Do While FileExist(strFolderPath & "\" & SafeFileName)
            lngCount = lngCount + 1
            SafeFileName = strFileName & "[" & CStr(lngCount) & "]" & strFileExt
        Loop

End Function

'creates a filename that doesnt exist by adding an number to filename, ex: filename[2].ext
'this function works with an entire complete filepath
Function SafeFilePath(ByVal strFilePath As String) As String
    'coded by Lewis Miller

      Dim strFileExt    As String
      Dim strNumber     As String
      Dim lngCount      As Long
      Dim lngPlace      As Long
      
    On Error Resume Next
        If InStr(strFilePath, ".") Then
            lngPlace = InStrRev(strFilePath, ".")
            strFileExt = Mid$(strFilePath, lngPlace)
            strFilePath = Left$(strFilePath, lngPlace - 1)
        End If
        Do While FileExist(strFilePath & strNumber & strFileExt)
            lngCount = lngCount + 1
            strNumber = "[" & CStr(lngCount) & "]"
        Loop
        SafeFilePath = strFilePath & strNumber & strFileExt
        
End Function


