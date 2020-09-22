Attribute VB_Name = "modBrowse"
Option Explicit
'=====================================================================================
' Browse for a Folder using SHBrowseForFolder API function with a callback
' function BrowseCallbackProc. Can also include files

'Original Code By:
' Stephen Fonnesbeck
' steev@xmission.com
' http://www.xmission.com/~steev
' Feb 20, 2000

'Modified By:
' Lewis Miller
' dethbomb@hotmail.com
'removed any unnecesary code and variables(const's etc)
'optimized all functions and added cotaskmemfree(),getactivewindow()
'removed lstrcat() - not needed

'Private Const BFFM_ENABLEOK As Long = 1125

Private Type BrowseInfo
    hwndOwner      As Long
    pIDLRoot       As Long
    pszDisplayName As String
    lpszTitle      As String
    ulFlags        As BROWSE_OPTIONS
    lpfnCallback   As Long
    lParam         As Long
    iImage         As Long
End Type

Private Type ITEMID
    cb As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As ITEMID
End Type

Public Enum CSIDL_INFO
    CSIDL_DESKTOP = &H0
    CSIDL_PROGRAMS = &H2
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_STARTMENU = &HB
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
End Enum

Public Enum BROWSE_OPTIONS
    BIF_NONE = &H0              'flag added to test for empty
    BIF_RETURNONLYFSDIRS = &H1
    BIF_DONTGOBELOWDOMAIN = &H2
    BIF_STATUSTEXT = &H4&
    BIF_RETURNFSANCESTORS = &H8
    BIF_EDITBOX = &H10
    BIF_VALIDATE = &H20
    BIF_NEWDIALOGSTYLE = &H40
    BIF_BROWSEINCLUDEURLS = &H80
    BIF_BROWSEFORCOMPUTER = &H1000
    BIF_BROWSEFORPRINTER = &H2000
    BIF_BROWSEINCLUDEFILES = &H4000
    BIF_SHAREABLE = &H8000
End Enum

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHBrowseForFolder Lib "SHELL32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "SHELL32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Any)
Private Declare Function GetActiveWindow Lib "user32" () As Long

Private strCurrentDirectory As String

Public Function SpecialFolder(CSIDL As CSIDL_INFO) As String
    On Error Resume Next
    Dim FolderPath As String * 260
    Dim IDL As ITEMIDLIST
    If SHGetSpecialFolderLocation(0&, CSIDL, IDL) = 0 Then
        If CStr(SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal FolderPath$)) Then
            SpecialFolder = Left$(FolderPath, InStr(FolderPath, vbNullChar) - 1)
        End If
        Call CoTaskMemFree(ByVal VarPtr(IDL))
    End If
End Function

Public Function ShowBrowse(Optional ByVal DialogTitle As String, Optional ByVal StartDir As String, Optional ByVal BrowseOptions As BROWSE_OPTIONS) As String

    Dim lpIDList As Long
    Dim tBrowseInfo As BrowseInfo

    If Len(StartDir) = 0 Then
        StartDir = CurDir$
    End If
    strCurrentDirectory = StartDir

    With tBrowseInfo
        .hwndOwner = GetActiveWindow
        .lpszTitle = DialogTitle
        If BrowseOptions = BIF_NONE Then
            'UNTESTED: BIF_NEWDIALOGSTYLE may not work on win95/98 (works on winME) it adds the "new folder" button
            .ulFlags = BIF_DONTGOBELOWDOMAIN Or BIF_STATUSTEXT Or BIF_NEWDIALOGSTYLE
        Else
            .ulFlags = BrowseOptions
        End If
        .lpfnCallback = FunctionAddress(AddressOf BrowseCallbackProc)
    End With

    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        ShowBrowse = Space$(260)
        If SHGetPathFromIDList(lpIDList, ShowBrowse) Then
            ShowBrowse = Left$(ShowBrowse, InStr(ShowBrowse, vbNullChar) - 1)
        Else
            ShowBrowse = ""
        End If
        Call CoTaskMemFree(lpIDList)
    End If

End Function

Private Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lp As Long, ByVal pData As Long) As Long

    Dim strBuffer As String

    Select Case uMsg
        Case 1       'BFFM_INITIALIZED
            Call SendMessage(hWnd, 1126, 1, ByVal strCurrentDirectory)

        Case 2       'BFFM_SELCHANGED
            strBuffer = Space$(260)
            If SHGetPathFromIDList(lp, strBuffer) Then
                strBuffer = Left$(strBuffer, InStr(strBuffer, vbNullChar) - 1)
                If strBuffer <> strCurrentDirectory Then
                    Call SendMessage(hWnd, 1124, 0, ByVal strBuffer)
                End If
            End If
        
        Case 3       'BFFM_VALIDATEFAILEDA
        Case 4       'BFFM_VALIDATEFAILEDW
    End Select

    BrowseCallbackProc = 0

End Function

' Assign a function pointer to a variable.
Private Function FunctionAddress(Address As Long) As Long

    FunctionAddress = Address

End Function

