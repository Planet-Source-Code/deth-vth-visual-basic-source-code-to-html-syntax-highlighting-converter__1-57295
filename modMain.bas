Attribute VB_Name = "modMain"
Option Explicit

'generic app starter module - by Lewis Miller

'main startup form, frmMain can then be accessed
'from anywhere in your project
Public frmMain As MainForm    '(As FORMNAME - the name of your main form ex: Form1)

'holds the path to app, available from anywhere in code
Public strAppPath As String

'public variable which is true when running in IDE
Public InIDE As Boolean


'Note: A proper application should use this method to start there
'program. This allows you to check for dependencies and initialize
'public variables. In you project properties set your 'startup object'
'to Sub Main(), this will cause vb to use this sub as the first starting
'point.

Sub Main()
    
    'This never happens in a compiled app,
    'thus InIde (in gen dec's) will be false.
    'Logic thanks to paul caton...
    Debug.Assert (SetRuntimeMode = False)

    'init the random number generator only once per app startup
    Randomize Timer

    'this avoids many calls to app.path...
    'some functions depend on this public variable
    strAppPath = App.Path
    If Left$(strAppPath, 2) = "\\" Then
        MsgBox "Please use local hardrive type fixed disk to run this program!", vbCritical, "Network Drive Error!"
        Exit Sub     '//quit
    End If


    '--------------------------------------------
    'do any other initialization here (example: loading settings, checking for ocx files etc...)

    'init ini module
    Ini_Initialize

    InitColorArray
    LoadColorSettings
    LoadKeywords

    '--------------------------------------------
    'process command line
    If ParseCommandLine(Command$) Then    'returns true if no gui
        Exit Sub
    End If


    '--------------------------------------------
    'load the startup form if any...
    Set frmMain = New MainForm    '(New FORMNAME - where FORMNAME is the name of your start form)

    'Note: Use frmMain as your main form variable instead of 'Form1'
    '(or whatever the name of your main form is).

    'load the form
    Load frmMain

    'center the form
    Center frmMain

    'show the form
    frmMain.Show

End Sub

'centers a form on the desktop
Sub Center(frmTarget As Form)

    Dim lngWidthMod As Long, lngHeightMod As Long

    With frmTarget
        'a little left of center, comment it out to use pure center
        If (Screen.Width - .Width) / 2 > 1000 Then
            lngWidthMod = 1000
        End If
        'a little above of center, comment it out to use pure center
        If (Screen.Height - .Height) / 2 > 600 Then
            lngHeightMod = 600
        End If

        'move form
        .Move ((Screen.Width - .Width) / 2) - lngWidthMod, ((Screen.Height - .Height) / 2) - lngHeightMod
    End With

End Sub

Function ParseCommandLine(ByVal Commandline As String) As Boolean

    'add your commandline parser here
    'and return true to skip showing a form
    'example : ParseCommandLine = true
    'example2: ParseCommandLine = ProcessCommandLine(Commandline)

    If Len(Commandline) > 0 Then
        blnCmdMode = True
        Commandline = Replace$(Commandline, vbQuote, "")
        If FileExist(Commandline) Then
            ProcessCommandLine Commandline
        Else
            MsgBox "Unable to load " & Commandline, vbCritical
        End If
        ParseCommandLine = True
    End If

End Function

Function ProcessCommandLine(ByVal Commandline As String) As Boolean

    'add your commandline processing code here
    'and return true to skip showing a form (or success)
    'example: ProcessCommandLine = true
    
    Dim strNewFile As String
        strNewFile = FileAddExtension(Commandline, "html")
        If ConvertFile(Commandline, strNewFile, True, True, False) Then
           MsgBox "Conversion successful! File saved to " & strNewFile, vbInformation
        Else
            MsgBox "Conversion was unsuccessful! Please check filepath and try again.", vbCritical
        End If

End Function

'this function is called from SubMain() to set a public
'variable to true if running in the visual basic IDE
'This variable (InIDE) can then be accessed anywhere in your project code
Function SetRuntimeMode() As Boolean
'logic thanks to paul caton
    InIDE = True
End Function
