Attribute VB_Name = "modCode"
Option Explicit

'---------------------------------------------------------------------------------------
' Module    : modCode
' DateTime  : 11/03/2004 22:06
' Author    : Lewis Miller
' Purpose   : Converting, parsing, tokenizing, and helper functions
'---------------------------------------------------------------------------------------
'this module is the "meat" of this vb to html syntax conversion program _
i also used this module as a test base for converting code


'("General Software Agreement")
'This code is provided "AS IS" without any warranty, express or implied, of any kind.
'The original software code author ("Lewis Miller") will accept no responsibility for
'any damages, monetary or otherwise, resulting from your use of this code. By your use
'of this code you agree to release the original software code author of and from any
'liabilty whatsoever resulting from any events that might occur related to this software
'code. You are free to use, distribute, and make copies of this code in anyway you wish
'so long as 1) You do not charge money or fees for this code, and 2) the original
'author's name remains intact. You may charge fees for a larger software project that
'this code is a part of, but not for this code as a sole item of sale. This code may not
'be sold for profit. 3) If you make any modifications to this code you MUST clearly note
'it as modified and changed by the original author of that additional software code
'before distributing it to others. 4) You MUST leave written software authoring credit
'intact within the source code to the original software code author. 5) This original
'software agreement will supercede any other new agreement here-in not originally written
'by the original software author, However the original software code author reserves the
'right to change this original agreement at any time or place without notice of any kind.
'While every effort has been made to provide software code free of bugs, you understand
'that due to the very nature of software and multiple, complex and different computer
'systems, software code can sometimes cause unexpected problems. Please notify the author
'if such an event should occur, he is always interested in paranormal phenomena.
'("End of General Software Agreement")


#If Win32 Then       'this is just for testing compiler directive's
    Public Declare Function timeGetTime Lib "winmm.dll" () As Long
#Else
    Public Declare Function timeGetTime Lib "winmm.dll" () As Integer
#End If

'since vb lacks this...
Public Const vbQuote   As String = """"    'same as Chr(34)

'this is the maximum number of characters the 'current' variable
'buffer in Tokenize() will hold before failing if it cant find a match.
'you can change it to less or more if you wish depending on your variable
'naming habits... if you always use short variable names use less, else more...
Private Const MAX_LEN  As Long = 50

Public Html                 As CStringList
Public ColorArr(8)          As Long      'colors used to colorize html
Public DefColArr(8)         As Long      'default colors
Public HexArr(8)            As String    'vb colors to html colors
Public KeyWords()           As String    'vb keywords
Public KeyWordCount         As Long      'number of keywords
Public ParseError           As Boolean   'parsing error flag
Public ErrorMessage         As String    'parsing error message
Public BlnNextLineIsComment As Boolean   '(next line of code is a comment)
Public blnAddLine           As Boolean   'add horizontal line flag
Public blnCmdMode           As Boolean   'command line mode

'variables accessed from modules and main form
'to keep status of conversion and progress
Public blnCancelled         As Boolean
Public StartTime            As Long
Public TotalLines           As Long
Public TotalFiles           As Long
Public TotalTime            As Long
Public ProjectName          As String


Public Enum TOKEN_TYPE 'TT
    TT_UNKNOWN
    TT_TAB           'tab
    TT_SPACE         'space
    TT_STRINGBEGIN   'double quote
    TT_STRINGTEXT    'string
    TT_STRINGEND     'double quote
    TT_LISTBEGIN     ' (
    TT_LISTSEPERATOR ' ,
    TT_LISTEND       ' )
    TT_LINENUM       'please, dont use them !
    TT_LINESTART     ' :
    TT_LINECONTINUE  ' _  '(hmm comments can continue...ugh (bug fix))
    TT_KEYWORD       'vb keyword
    TT_OPERATOR      'vb operator + = - / \ > < <> <= >= ^ & *
    TT_SCOPEDCALL    ' inside a With (.Function), or "object.function" or a dot inside a "vari.able"
    TT_VARIABLE      'any text not a keyword, number, or not anything else in this list
    TT_HEXOROCTAL    '&H0FFF &O10
    TT_VARTYPE       ' 999# 999$ 999% 999& 999.00! etc...
    TT_NEGATIVE      ' -1  -(varX) -varX
    TT_NUMERIC       'any number
    TT_COMMENTSTART  '<-- that
    TT_COMMENTTEXT   ' this
    TT_FORCEVARBEGIN ' [  - as in mCol.[_NewEnum]
    TT_FORCEVAREND   ' ]
    TT_DIRECTIVE     ' #Const #If #Else etc
    TT_EOL           '(crlf) end of line
    TT_DIVIDER       'sub/function/property  divider line added by vth
End Enum

'the 8 types of colors used
Public Enum COLOR_INDEX 'COL
    COL_COMMENT = 0
    COL_KEYWORD = 1
    COL_IDENTIFIER = 2
    COL_STRING = 3
    COL_NORMAL = 4
    COL_OPERATOR = 5
    COL_NUMBER = 6
    COL_DIRECTIVE = 7
    COL_BACKGROUND = 8
End Enum

'type used to hold project file values
'used to create a master index page of a project
Type VB_Project
    Initialized      As Boolean
    BaseDir          As String
    SaveDir          As String
    Properties       As New Collection
    References       As New Collection
    Other            As New Collection
    Files            As New Collection
End Type

Public Project As VB_Project

'incomplete... but works.
'this create a html index links page in a converted project folder
Sub CreateMasterIndexPage(ByVal strMasterIndexName As String, ByVal blnOpenIndex As Boolean)

    Dim FirstPart       As String
    Dim LastPart        As String
    Dim strHtmlPage     As String
    Dim lngSplitStart   As Long
    Dim vntItem         As Variant
    Dim intFileNum      As Integer
    Dim strAppVer       As String
    Dim strArr()        As String

    With Project     'i love 'With'...saves typing :)

        If Not .Initialized Then
            MsgBox "Project is not initialized or loaded, cannot create index page!", vbCritical
            Exit Sub
        End If

        'get html page template from res file
        strHtmlPage = StrConv(LoadResData("index", "html"), vbUnicode)
        'seperate it
        lngSplitStart = InStr(strHtmlPage, "########")
        FirstPart = Left$(strHtmlPage, lngSplitStart - 1)
        LastPart = Mid$(strHtmlPage, lngSplitStart + 8)

        On Error Resume Next    'this is in case a item we want isnt in the project 'other' collection
        'compile version number
        strAppVer = .Other("MajorVer") & "." & .Other("MinorVer") & "." & .Other("RevisionVer")

        'replace stuff in the template
        FirstPart = Replace$(FirstPart, "PROJECT_TITLE", .Other("Title"))
        FirstPart = Replace$(FirstPart, "PROJECT_NAME", .Other("Name"))
        If .Other("VersionCompanyName") = "" Then
            FirstPart = Replace$(FirstPart, "By APPLICATION_AUTHOR", "")
            FirstPart = Replace$(FirstPart, "APPLICATION_AUTHOR", " ")
        Else
            FirstPart = Replace$(FirstPart, "APPLICATION_AUTHOR", .Other("VersionCompanyName"))
        End If
        FirstPart = Replace$(FirstPart, "APPLICATION_NAME", .Other("VersionProductName"))
        FirstPart = Replace$(FirstPart, "APPLICATION_TITLE", .Other("Title") & " " & strAppVer & " ")
        FirstPart = Replace$(FirstPart, "APPLICATION_VERSION", strAppVer)

        FirstPart = FirstPart & vbCrLf & "<br><font size=" & vbQuote & "2" & vbQuote & " face=" & vbQuote & "sans-serif" & vbQuote & " color=" & vbQuote & "black" & vbQuote & ">" & .Other("Name") & " contains " & CStr(.Other("FileCount")) & " code files with a total of " & CStr(.Other("LineCount")) & " lines of code.<br><br>" & vbCrLf
        On Error GoTo 0

        'write out references
        If .References.Count > 0 Then
            FirstPart = FirstPart & vbCrLf & "<font size=" & vbQuote & "2" & vbQuote & " face=" & vbQuote & "sans-serif" & vbQuote & " color=" & vbQuote & "black" & vbQuote & ">This project contains the following " & CStr(.References.Count) & " references and objects<ul>" & vbCrLf
            For Each vntItem In .References
                FirstPart = FirstPart & "<li>" & vntItem & "</li>" & vbCrLf
            Next
            FirstPart = FirstPart & "</ul></font>" & vbCrLf
        End If

        If .Files.Count > 0 Then
            'write out links to code pages
            FirstPart = FirstPart & "<font face=" & vbQuote & "sans-serif" & vbQuote & "><b>Click A Button Below To View The Source Code</b></font><br>" & vbCrLf
            For Each vntItem In .Files
                strArr = Split(vntItem, "#")
                FirstPart = FirstPart & HTMLButton("<a href=" & vbQuote & strArr(1) & vbQuote & ">" & strArr(0) & "</a>" & vbCrLf)
            Next
        End If

        'create the extra properties page
        If .Properties.Count > 0 Then
            FirstPart = FirstPart & HTMLButton("<a href=" & vbQuote & "project.html" & vbQuote & ">More Project Info</a>" & vbCrLf)
            intFileNum = FreeFile
            Open .SaveDir & "\project.html" For Output As intFileNum
            Print #intFileNum, "<html><head><title>More Project Info</title></head><body>" & vbCrLf & "<font size=4 color=#0000CC>Miscellaneous Project Properties&nbsp;&nbsp;</font><a href=" & vbQuote & strMasterIndexName & vbQuote & ">Go Back</a><br>" & vbCrLf & "<font size=2 color=#0000FF><br>" & vbCrLf
            For Each vntItem In .Properties
                If Len(vntItem) > 0 Then
                    Print #intFileNum, vntItem & "<br>" & vbCrLf
                End If
            Next
            Print #intFileNum, "<br></font><a href=" & vbQuote & strMasterIndexName & vbQuote & ">Go Back</a><br><br></body></html>" & vbCrLf
            Close intFileNum
        End If

        'create the form buttons pic
        intFileNum = FreeFile
        If FileExist(.SaveDir & "\buttons.gif") Then FileKill .SaveDir & "\buttons.gif"
        Open .SaveDir & "\buttons.gif" For Binary As intFileNum
        Put #intFileNum, , CStr(StrConv(LoadResData("buttons", "pic"), vbUnicode))
        Close intFileNum

        'eventually this will extract the actual icon from frx or exe file, for now we use default icon
        intFileNum = FreeFile
        If FileExist(.SaveDir & "\icon.gif") Then FileKill .SaveDir & "\icon.gif"
        Open .SaveDir & "\icon.gif" For Binary As intFileNum
        Put #intFileNum, , CStr(StrConv(LoadResData("form", "pic"), vbUnicode))
        Close intFileNum

        'check to see if index page already exists
        If FileExist(.SaveDir & "\" & strMasterIndexName) Then FileKill .SaveDir & "\" & strMasterIndexName
        intFileNum = FreeFile

        'we have to error trap this because it may be a bogus user-input file name
        On Error Resume Next
        Open .SaveDir & "\" & strMasterIndexName For Binary As intFileNum
        If Err Then
            Close #intFileNum
            MsgBox "A fatal error has occured while trying to write to the file " & strMasterIndexName & ". Please check the filename and try again. [ debug: " & Err.Description & " ]", vbCritical
            Exit Sub
        Else
            Put #intFileNum, , FirstPart & "<br>" & LastPart
        End If
        Close intFileNum
        On Error GoTo 0

        'open it or not?
        If blnOpenIndex Then    'open it
            ExecuteFile .SaveDir & "\" & strMasterIndexName
        Else
            MsgBox "Master index page (" & .SaveDir & "\" & strMasterIndexName & ") created.", vbInformation
        End If

    End With

End Sub

'creates a html vb-like button (well, close enough...)
Function HTMLButton(ByVal strButtonText As String) As String
    HTMLButton = "<br><table width=" & vbQuote & "25%" & vbQuote & " border=" & vbQuote & "1" & vbQuote & " cellspacing=" & vbQuote & "2" & vbQuote & " cellpadding=" & vbQuote & "2" & vbQuote & " bgcolor=" & vbQuote & "#D5D0BF" & vbQuote & " bordercolor=" & vbQuote & "#999999" & vbQuote & " align=" & vbQuote & "center" & vbQuote & "> <tr><td bgcolor=" & vbQuote & "#D5D0BF" & vbQuote & " bordercolor=" & vbQuote & "#CCCCCC" & vbQuote & " nowrap><center>" & strButtonText & "</center></td></tr></table>" & vbCrLf
End Function

'wraps font color html around a string
Function WrapColor(ByVal ColorIndex As COLOR_INDEX, ByVal strToWrap As String) As String
    WrapColor = "<font color=" & HexArr(ColorIndex) & ">" & strToWrap & "</font>"
End Function

'error trapped function to test if an item has been added to Project UDT
Function PropertyExist(ByVal strPropName As String) As Boolean
    Dim strTest As String
    
On Error GoTo PropertyExist_Error

    If Project.Initialized Then
            strTest = Project.Other(strPropName)
            PropertyExist = True
    End If

   On Error GoTo 0
   Exit Function

PropertyExist_Error:

End Function

'add a property from a project file to the Project UDT
Sub AddProjectProperty(ByVal strProperty As String)

    Dim PropArr() As String
    Dim ReferenceArr() As String
    
    On Error Resume Next 'just in case
    
    With Project
        If InStr(strProperty, "=") Then
            PropArr = Split(strProperty, "=", 2)

            Select Case PropArr(0)
                Case "Reference"
                    If InStr(PropArr(1), "#") Then
                        ReferenceArr = Split(PropArr(1), "#")
                    Else
                        ReDim ReferenceArr(4) As String
                        ReferenceArr(3) = PropArr(1)
                    End If

                    If UBound(ReferenceArr) = 4 Then
                        .References.Add GetFileName(ReferenceArr(3)) & vbTab & ":" & vbTab & ReferenceArr(0)
                    End If

                Case "Object"

                    If InStr(PropArr(1), "#") Then
                        ReferenceArr = Split(PropArr(1), "#")
                    Else
                        ReDim ReferenceArr(2) As String
                        ReferenceArr(2) = PropArr(1)
                    End If

                    If UBound(ReferenceArr) = 2 Then
                        .References.Add GetFileName(Mid$(ReferenceArr(2), InStr(ReferenceArr(2), ";") + 2)) & vbTab & " : " & vbTab & ReferenceArr(0)
                    End If

                Case "Form", "Class", "Module", "PropertyPage", "UserControl", "Designer"
                        'do nothing, but keeps this from going to 'case else'
                
                Case Else
                    .Other.Add Replace$(PropArr(1), """", ""), PropArr(0)
                    .Properties.Add PropArr(0) & vbTab & "=" & vbTab & Replace$(PropArr(1), """", "")
            
            End Select
        End If
    End With
   
   On Error GoTo 0

End Sub


'This function inserts strToInsert into strDestination starting at lngInsertPlace.
'If lngInsertPlace is 1 (or zero) strToInsert is placed at the beginning.
'If lngInsertPlace is equal to (or greater than) strDestination's length it is placed at the end.
'Otherwise everything from lngInsertPlace to the end of strDestination is scooted forward
'to make room for strToInsert, and strToInsert is placed at lngInsertPlace.
'This function was coded with large strings in mind...
Function Insert(ByVal lngInsertPlace As Long, ByVal strToInsert As String, ByVal strDestination As String) As String
    'Written by Lewis Miller

    Dim lngDestLen As Long
    Dim lngSrcLen As Long

    lngDestLen = Len(strDestination)
    If lngDestLen = 0 Then    'no dest str
        Insert = strToInsert
    Else
        lngSrcLen = Len(strToInsert)
        If lngSrcLen = 0 Then    'no insert str
            Insert = strDestination
        Else
            Insert = Space$(lngDestLen + lngSrcLen)
            If lngInsertPlace < 2 Then    'insert at beginning
                Mid(Insert, 1, lngSrcLen) = strToInsert
                Mid(Insert, lngSrcLen + 1, lngDestLen) = strDestination
            ElseIf lngInsertPlace >= lngDestLen Then    'insert at end
                Mid(Insert, 1, lngDestLen) = strDestination
                Mid(Insert, lngDestLen + 1, lngSrcLen) = strToInsert
            Else     'insert into middle somewhere
                Mid(Insert, 1, lngInsertPlace - 1) = Left$(strDestination, lngInsertPlace - 1)
                Mid(Insert, lngInsertPlace, lngSrcLen) = strToInsert
                Mid(Insert, lngInsertPlace + lngSrcLen, (lngDestLen + 1) - lngInsertPlace) = Mid$(strDestination, lngInsertPlace)
            End If
        End If
    End If

End Function

'repeat a string 'lngRepeatCount' number of times
Function Repeat(ByVal strRepeatItem As String, ByVal lngRepeatCount As Long) As String
    'Written by Lewis Miller

    Dim lngLength As Long
    Dim lngIndex As Long

    lngLength = Len(strRepeatItem)
    If (lngLength > 0) And (lngRepeatCount > 0) Then
        Repeat = Space$(lngLength * lngRepeatCount)
        For lngIndex = 0 To lngRepeatCount - 1
            Mid$(Repeat, ((lngIndex * lngLength) + 1), lngLength) = strRepeatItem
        Next lngIndex
    End If

End Function

'default colors
Public Sub InitColorArray()
    DefColArr(0) = 32768
    DefColArr(1) = 8388608
    DefColArr(3) = 8388608
    DefColArr(7) = 4210752
    DefColArr(8) = 16777215
    'the rest are default 0 (black)
End Sub

Sub LoadColorSettings()

    Dim lngLoopIndex As Long

    For lngLoopIndex = 0 To 8
        ColorArr(lngLoopIndex) = ReadNumber("colors", CStr(lngLoopIndex), DefColArr(lngLoopIndex))
        HexArr(lngLoopIndex) = vbQuote & HtmlColor(ColorArr(lngLoopIndex)) & vbQuote
    Next lngLoopIndex

    blnAddLine = CBool(ReadNumber("options", "addline", 1))

End Sub

Sub SaveColorSettings()

    Dim lngLoopIndex As Long

    For lngLoopIndex = 0 To 8
        WriteNumber "colors", CStr(lngLoopIndex), ColorArr(lngLoopIndex)
    Next lngLoopIndex

    WriteNumber "options", "addline", Abs(CInt(blnAddLine))
End Sub

'load keywords from res file
'todo: provide option to load from text file
Sub LoadKeywords()

    'load from res file
    KeyWords = Split(StrConv(LoadResData("keywords", "text"), vbUnicode), vbCrLf)

    'remove any empty ones at the end of array
    Do While UBound(KeyWords) > 0
        If Len(KeyWords(UBound(KeyWords))) = 0 Then
            ReDim Preserve KeyWords(UBound(KeyWords) - 1) As String
        Else
            Exit Do
        End If
    Loop
    
    KeyWordCount = UBound(KeyWords) + 1
    
    'Note: you can save the keywords to file with the following code:
    'FileSave CStr(StrConv(LoadResData("keywords", "text")), vbUnicode), strAppPath & "\keywords.txt"
    
End Sub


'change a numeric color to a string html color
Function HtmlColor(ByVal Color As Long) As String

    Dim lngRed As Long, lngGreen As Long, lngBlue As Long
    lngRed = Color Mod 256&
    Color = Color \ 256&
    lngGreen = Color Mod 256&
    Color = Color \ 256&
    lngBlue = Color Mod 256&
    HtmlColor = "#" & Right$("00" & Hex$(lngRed), 2) & Right$("00" & Hex$(lngGreen), 2) & Right$("00" & Hex$(lngBlue), 2)

End Function


'check to make sure that the progress page is saved to disk
'and returns the path to it
Function ProgressPagePath() As String

    Dim strProgBarPath As String

    strProgBarPath = strAppPath & "\progbar.gif"
    ProgressPagePath = strAppPath & "\progress.html"

    'does the progbar pic exist?
    If Not FileExist(strProgBarPath) Then
        'no, so save it from res file
        FileByte LoadResData("progbar", "pic"), strProgBarPath
    End If

    'does the progress html page exist?
    If Not FileExist(ProgressPagePath) Then
        'no, so save it from res file
        FileByte LoadResData("progress", "html"), ProgressPagePath
    End If

End Function

'checks to make sure the error page is saved to disk and
'returns a path to it
Function ErrorPage() As String

    ErrorPage = strAppPath & "\error.html"

    If Not FileExist(ErrorPage) Then
        FileWrite "<html><head><title>Error!</title></head><body><font color=red><i>An Error has Occurred While Parsing Code.</i></font></body></html>", ErrorPage
    End If

End Function

'checks to make sure the start page is saved to disk and
'returns a path to it
Function StartPage(imgSave As Image) As String

    StartPage = strAppPath & "\welcome.html"

    If Not FileExist(StartPage) Then
        FileWrite "<html><head><title>Welcome to VHT!</title></head><body><font color=red size=5><i>Welcome To VTH.<br></i></font><font color=green size=2>'The VB Syntax To Html Wizard</font><br><br>" & Repeat("&nbsp;", 15) & "<img src=icon.ico></img></body></html>", StartPage
        On Error Resume Next
        SavePicture imgSave.Picture, strAppPath & "\icon.ico"
        On Error GoTo 0
    End If

End Function

'checks to make sure the blank page is saved to disk and
'returns a path to it
Function BlankPage() As String

    BlankPage = strAppPath & "\blank.html"

    If Not FileExist(BlankPage) Then
        FileWrite "<html><head><title>Blank</title></head><body></body></html>", BlankPage
    End If

End Function

'saves the html source list to a local page for previewing
Function MakeHtmlPage(Optional ByVal strFilePath As String) As String

    Dim lngLoopIndex As Long, intFileNum As Integer

    If Len(strFilePath) > 0 Then
        MakeHtmlPage = strFilePath
    Else
        MakeHtmlPage = strAppPath & "\preview.html"
    End If

   On Error GoTo Failure

    If Not Html Is Nothing Then
        With Html
            If .ListCount > 0 Then
                If FileExist(MakeHtmlPage) Then
                    FileKill MakeHtmlPage
                End If
                intFileNum = FreeFile
                Open MakeHtmlPage For Binary As #intFileNum
                Put #intFileNum, , "<html><head><title>Source Code Preview</title></head><body>" & vbCrLf
                For lngLoopIndex = 0 To .ListCount - 1
                    Put #intFileNum, , .List(lngLoopIndex)
                Next lngLoopIndex
                Put #intFileNum, , "</body></html>" & vbCrLf
                Close #intFileNum
                Exit Function
            End If
        End With
    End If

Failure:
    FileWrite "<html><head><title>Error!</title></head><body><font color=red><i>An Error Has Occurred While Creating An Html Preview Of Converted Source Code.</i></font></body></html>", MakeHtmlPage

End Function


Function GetVBPath(ByVal TempPath As String) As String

    Dim lngPlace As Long
    Dim strType As String
    
    'Class=Class1; Class1.cls
    lngPlace = InStr(TempPath, "=")
    If lngPlace Then
        strType = Left$(TempPath, lngPlace - 1)
        TempPath = Trim$(Mid$(TempPath, lngPlace + 1))
    End If

    'Class1; Class1.cls
    lngPlace = InStr(TempPath, ";")
    If lngPlace Then
        TempPath = Trim$(Mid$(TempPath, lngPlace + 2))
    Else
        TempPath = Trim$(TempPath)
    End If
    
    GetVBPath = TempPath

End Function

'the entire reason for these next two functions is that when changing extensions
'on vb files to .html, it can create ambiguous file names, like for
'example : sys.bas, sys.ctl, sys.frm, would all be named sys.html
'By using the base filename as a key, when adding it to a collection, it will
'error if it already exists, so we append a [<Num>] to base file name
'till it succeeds....
Function CreateHtmlPath(ByVal strFile As String)

     Dim lngCount   As Long
     Dim lngPlace   As Long
     Dim strNewFile         As String
     Dim strOriginalBase    As String
     Dim strNewExtension    As String
     Dim strOldExtension    As String
     
     lngPlace = InStrRev(strFile, ".")
     
     strNewExtension = ".html"
     strOriginalBase = Left$(strFile, lngPlace - 1)
     strOldExtension = Mid$(strFile, lngPlace + 1)
     strNewFile = strOriginalBase
     
     Do While Not AddHtmlFile(strOriginalBase & " (" & strOldExtension & ")#" & strNewFile & strNewExtension, strNewFile)
       lngCount = lngCount + 1
       strNewFile = strOriginalBase & "[" & CStr(lngCount) & "]"
     Loop
    
    CreateHtmlPath = strNewFile & strNewExtension
    
End Function

'false if exists, else added ok
Function AddHtmlFile(ByVal strValName As String, ByVal strKeyName As String) As Boolean
      
      On Error Resume Next
      Project.Files.Add strValName, strKeyName
      AddHtmlFile = (Err = 0)
      On Error GoTo 0
      
End Function

'checks to see if a file has a html header
Function FileHasHeader(ByVal strFilePath As String) As Boolean

    Dim strFileHeader As String
    Dim intFileNum As Integer
    Dim lngAllocLength As Long

    If FileExist(strFilePath) Then
        lngAllocLength = GetFileLength(strFilePath)
        If lngAllocLength > 512 Then
            lngAllocLength = 512
        End If

        intFileNum = FreeFile
        On Error Resume Next
        Open strFilePath For Binary As #intFileNum
        If Err = 0 Then
            strFileHeader = Space$(lngAllocLength)
            Get #intFileNum, , strFileHeader

            FileHasHeader = (InStr(1, strFileHeader, "<html>", vbTextCompare) > 0)
        End If
        Close intFileNum
        On Error GoTo 0
    End If

End Function

'prepare html list for new code
Sub PrepStack()

    Set Html = New CStringList
  With Html
    .AddItem "<table width=" & vbQuote & "95%" & vbQuote & " border=" & vbQuote & "0" & vbQuote & " cellspacing=" & vbQuote & "0" & vbQuote & " cellpadding=" & vbQuote & "0" & vbQuote & " bgcolor=" & HexArr(COL_BACKGROUND) & " align=" & vbQuote & "center" & vbQuote & ">"
    .AddItem "  <tr>"    'new column
    .AddItem "   <td nowrap>"    'new row
    .AddItem "    <code>"    'code font
    .AddItem "     <font color=" & HexArr(COL_NORMAL) & ">"    'normal color
  End With
  
  'note: alternatively you can remove the <code> tags and change the last line to:
  '<font color=" & HexArr(COL_NORMAL) & " face=" & vbquote & "Courier New" & vbquote & ">"
  'you will have to remove </code> tag from CloseStack() function also
End Sub

'tie off html list loose ends
Sub CloseStack()

    If Not Html Is Nothing Then
        With Html
            .AddItem "    </font>"    'end normal color
            .AddItem "   </code>"    'end code font
            .AddItem "  </td>"    'end column
            .AddItem " </tr>"    'end row
            .AddItem "</table>"    'end table
        End With
    End If

End Sub

Function ConvertTokenListToHtml(Tokens As TokenList) As String
    'Written by Lewis Miller

    Dim lngCurrIndex  As Long

    If Not (Tokens Is Nothing) Then
        If Tokens.Count > 0 Then
            For lngCurrIndex = 1 To Tokens.Count
                With Tokens(lngCurrIndex)
                    If .Length Then
                        Select Case .Kind
                            Case TT_COMMENTSTART, TT_COMMENTTEXT
                                ConvertTokenListToHtml = ConvertTokenListToHtml & WrapColor(COL_COMMENT, SafeHtml(.Value))

                            Case TT_DIRECTIVE
                                ConvertTokenListToHtml = ConvertTokenListToHtml & WrapColor(COL_DIRECTIVE, SafeHtml(.Value))

                            Case TT_DIVIDER
                                ConvertTokenListToHtml = ConvertTokenListToHtml & vbCrLf & "<hr size=" & vbQuote & "2" & vbQuote & " align=" & vbQuote & "left" & vbQuote & "><br>" & vbCrLf

                            Case TT_EOL
                                ConvertTokenListToHtml = ConvertTokenListToHtml & "<br>"

                            Case TT_HEXOROCTAL, TT_LINENUM, TT_NEGATIVE, TT_NUMERIC
                                ConvertTokenListToHtml = ConvertTokenListToHtml & WrapColor(COL_NUMBER, SafeHtml(.Value))

                            Case TT_KEYWORD
                                ConvertTokenListToHtml = ConvertTokenListToHtml & WrapColor(COL_KEYWORD, SafeHtml(.Value))

                            Case TT_OPERATOR, TT_LINESTART
                                ConvertTokenListToHtml = ConvertTokenListToHtml & WrapColor(COL_OPERATOR, SafeHtml(.Value))

                            Case TT_STRINGBEGIN, TT_STRINGEND 'for empty strings ex: ""
                                If NextToken(lngCurrIndex, Tokens).Kind = TT_STRINGEND Then
                                  ConvertTokenListToHtml = ConvertTokenListToHtml & WrapColor(COL_STRING, vbQuote & vbQuote)
                                  lngCurrIndex = lngCurrIndex + 1
                                End If
                                
                            Case TT_STRINGTEXT
                                ConvertTokenListToHtml = ConvertTokenListToHtml & WrapColor(COL_STRING, vbQuote & SafeHtml(.Value) & vbQuote)

                            Case TT_VARIABLE
                                ConvertTokenListToHtml = ConvertTokenListToHtml & WrapColor(COL_IDENTIFIER, SafeHtml(.Value))

                            Case TT_SPACE
                                ConvertTokenListToHtml = ConvertTokenListToHtml & "&nbsp;"

                            Case TT_TAB
                                ConvertTokenListToHtml = ConvertTokenListToHtml & "&nbsp;&nbsp;&nbsp;&nbsp;"

                            Case TT_SCOPEDCALL
                                Select Case .Value
                                    Case "Debug.Print", "Debug.Assert"
                                        ConvertTokenListToHtml = ConvertTokenListToHtml & WrapColor(COL_KEYWORD, SafeHtml(.Value))

                                    Case Else
                                        ConvertTokenListToHtml = ConvertTokenListToHtml & SafeHtml(.Value)
                                End Select

                            Case Else
                                ConvertTokenListToHtml = ConvertTokenListToHtml & SafeHtml(.Value)
                          
                          End Select
                    End If
                End With
            Next lngCurrIndex
        End If
    End If


End Function

'html forces us to have to do this...
'As i quickly found out, having html in the source code
'itself (like this function) is a PITA!
Function SafeHtml(ByVal strUnsafeText As String) As String
    'Written by Lewis Miller

    'the original...
    '    strUnsafeText = Replace$(strUnsafeText, "&", "&amp;")   '(&) has to come first so you dont replace it in the others!
    '    strUnsafeText = Replace$(strUnsafeText, "  ", " &nbsp;")
    '    strUnsafeText = Replace$(strUnsafeText, "<", "&lt;")
    '    strUnsafeText = Replace$(strUnsafeText, ">", "&gt;")
    '    SafeHtml = Replace$(strUnsafeText, vbQuote, "&quot;")

    '&apos; is not needed...
    'strUnsafeText = Replace$(strUnsafeText, "'", "&apos;")

    'the new... faster because it forces vb to use its own stack variables for temporaries
    SafeHtml = Replace$(Replace$(Replace$(Replace$(Replace$(strUnsafeText, "&", "&amp;"), "  ", " &nbsp;"), "<", "&lt;"), ">", "&gt;"), vbQuote, "&quot;")
    'mmmmm... now thats nesting functions baby...

End Function

'safely grab the next character from a line of code relative to current position
Function NextChar(ByVal lngCurrentPos As Long, ByVal strCurBuffer As String) As String

    Dim lngBufferLength As Long

    lngBufferLength = Len(strCurBuffer)
    If lngBufferLength > 1 Then
        If lngCurrentPos < lngBufferLength And lngCurrentPos + 1 > 0 Then
            NextChar = Mid$(strCurBuffer, lngCurrentPos + 1, 1)
        End If
    End If

End Function

'safely grab the previous character from a line of code relative to current position
Function PrevChar(ByVal lngCurrentPos As String, ByVal strCurBuffer As String) As String

    Dim Length As Long

    Length = Len(strCurBuffer)
    If Length > 1 And lngCurrentPos > 1 Then
        If lngCurrentPos <= Length + 1 Then
            PrevChar = Mid$(strCurBuffer, lngCurrentPos - 1, 1)
        End If
    End If

End Function

'grab the last token from a list of tokens
Function LastToken(Tokens As TokenList, Optional ByVal blnFailOnEmpty As Boolean) As TokenItem

    If Not Tokens Is Nothing Then
        If Tokens.Count > 0 Then
            Set LastToken = Tokens(Tokens.Count)
            Exit Function
        End If
    End If
    
    If Not blnFailOnEmpty Then
        'in case of failure, provide a default empty one
        Set LastToken = New TokenItem
    End If

End Function

'get the next token past index, from a list of tokens
Function NextToken(ByVal lngCurrentTokenIndex As Long, Tokens As TokenList, Optional ByVal blnFailOnEmpty As Boolean) As TokenItem

    If Not Tokens Is Nothing Then
        If Tokens.Count > lngCurrentTokenIndex + 1 Then
            Set NextToken = Tokens(lngCurrentTokenIndex + 1)
            Exit Function
        End If
    End If
    
    If Not blnFailOnEmpty Then
        'in case of failure provide a default
        Set NextToken = New TokenItem
    End If

End Function

'test a string to see if it is a visual basic keyword
Function isKeyword(ByVal strTest As String) As Boolean

    Dim lngLoopIndex As Long
    
    For lngLoopIndex = 0 To KeyWordCount - 1    'keywordcount is a public variable
        If strTest = KeyWords(lngLoopIndex) Then
            isKeyword = True
            Exit Function
        End If
    Next lngLoopIndex

End Function '

'*****************************************************************************************
'   ***  This function is the MAIN TOKENIZER ENGINE. Close to 500 lines of code  ***
'it seperates a line of code into a list of tokens and decides what kind each token is.
'It is basically a loop with a large select case statement that contains nested if's...
'Uses lots of helper functions to keep the logic simpler.
'Because this is a beta application, in order to better catch bugs, there is little to
'no error handling as of yet.  If you find one please let me know.
'Be careful if you change something :)
'*****************************************************************************************
Function Tokenize(ByVal strCodeLine As String, ByVal lngLineNum As Long) As TokenList
    'Written by Lewis Miller

    Dim Token               As TokenItem    'used for look ahead and looping thru token list
    Dim strNextChar         As String    'stores next char
    Dim strPrevChar         As String    'stores previous char
    Dim strCurrentChar      As String    'stores current char
    Dim strCurrentBuffer    As String    'stores variable/keyword buffer
    Dim lngCurrentIndex     As Long    'current position
    Dim lngCodeLength       As Long    'length of code line
    Dim lngBufferLen        As Long    'length of strCurrentBuffer
    Dim lngBufferStart      As Long    'stores starting place of strCurrentBuffer within strCodeLine
    Dim lngNextPlace        As Long    'used for look ahead searches
    Dim RecurseList         As TokenList    'used for recursing

    'make sure error is cleared
    ErrorMessage = ""
    ParseError = False

    'examine global flag
    If BlnNextLineIsComment Then
        'add a comment character
        strCodeLine = Insert(0, "'", strCodeLine)
    End If

    'initialize
    Set Tokenize = New TokenList
    lngCodeLength = Len(strCodeLine)
    lngCurrentIndex = 1
    lngBufferStart = 1

    With Tokenize

        If strCodeLine = vbCrLf Then    'nothing to do
            .Add MakeToken(vbCrLf, TT_EOL)
            Exit Function
        End If


        'start looping through code line one char at a time
        Do While lngCurrentIndex <= lngCodeLength
            strCurrentChar = Mid$(strCodeLine, lngCurrentIndex, 1)

            'examine each character
            Select Case strCurrentChar

                Case "<", ">"    'possible double char operator
                    If Len(strCurrentBuffer) = 0 Then
                        strPrevChar = PrevChar(lngCurrentIndex, strCodeLine)
                        If strPrevChar <> "" And strPrevChar <> vbTab And strPrevChar <> " " Then
                            ParseError = True
                            ErrorMessage = "A syntax or parse error has occurred on line #" & CStr(lngLineNum) & ". An invalid operator ( " & strCurrentChar & " ) was found with incorrect spacing. Please check your code and try again." & vbCrLf & vbCrLf & "[ debug: " & Left$(strCodeLine, Len(strCodeLine) - 2) & " ]"
                            Exit Function
                        End If
                        strNextChar = NextChar(lngCurrentIndex, strCodeLine)
                        If strNextChar = "=" Or strNextChar = ">" Then    'double char operators   ( <= >= <> )
                            .Add MakeToken(strCurrentChar & strNextChar, TT_OPERATOR)
                            lngCurrentIndex = lngCurrentIndex + 1
                        ElseIf strNextChar = " " Then    'single char operators
                            .Add MakeToken(strCurrentChar, TT_OPERATOR)
                        Else    'syntax error
                            ParseError = True
                            ErrorMessage = "A syntax or parse error has occurred on line #" & CStr(lngLineNum) & ". An invalid operator ( " & strCurrentChar & " ) was found with incorrect spacing. Please check your code and try again." & vbCrLf & vbCrLf & "[ debug: " & Left$(strCodeLine, Len(strCodeLine) - 2) & " ]"
                            Exit Function
                        End If
                        'we know the last one is a space so add it, save a parse loop
                        lngCurrentIndex = lngCurrentIndex + 1
                        .Add MakeToken(" ", TT_SPACE)

                    Else    'syntax error
                        ParseError = True
                        ErrorMessage = "A syntax or parse error has occurred on line #" & CStr(lngLineNum) & ". An invalid operator ( " & strCurrentChar & " ) was found with incorrect spacing. Please check your code and try again." & vbCrLf & vbCrLf & "[ debug: " & Left$(strCodeLine, Len(strCodeLine) - 2) & " ]"
                        Exit Function
                    End If
                    lngCurrentIndex = lngCurrentIndex + 1
                    lngBufferStart = lngCurrentIndex


                Case "!", "#", "$", "%"    'type declarators
                    'todo: rework logic, may fail in unknown code
                    If Len(strCurrentBuffer) > 0 Then
                        .Add ResolveToken(strCurrentBuffer)
                        strCurrentBuffer = ""
                        .Add MakeToken(strCurrentChar, TT_VARTYPE)
                        lngCurrentIndex = lngCurrentIndex + 1
                    Else
                        If (strCurrentChar = "#") Then
                            'we need to figure out if this is a compiler directive, file number, or date/time literal else theres a syntax error
                            If Not IsFirstChar("#", strCodeLine, lngCodeLength) Then    'its not a compiler directive
                                'here we need to figure out if it is a datetime literal
                                ' #1/2/2003 2:30:00 PM#
                                lngNextPlace = InStr(lngCurrentIndex + 1, strCodeLine, "#")
                                If lngNextPlace > lngCurrentIndex + 1 Then
                                    strCurrentBuffer = Mid$(strCodeLine, lngCurrentIndex + 1, lngNextPlace - (lngCurrentIndex + 1))
                                    If IsDate(strCurrentBuffer) Then    'its a date, honey...
                                        .Add MakeToken("#", TT_VARTYPE)
                                        .Add MakeToken(strCurrentBuffer, TT_NUMERIC)
                                        lngCurrentIndex = lngNextPlace
                                    End If
                                End If
                                .Add MakeToken("#", TT_VARTYPE)
                                lngCurrentIndex = lngCurrentIndex + 1

                            Else
                                'if the next char is alphabetical  then its probably a compiler directive
                                lngBufferStart = lngCurrentIndex
                                lngCurrentIndex = lngCurrentIndex + 1
                                strCurrentChar = Mid$(strCodeLine, lngCurrentIndex, 1)
                                Do While IsAlpha(strCurrentChar) And lngCurrentIndex < lngCodeLength
                                    strCurrentBuffer = strCurrentBuffer & strCurrentChar
                                    lngCurrentIndex = lngCurrentIndex + 1
                                    strCurrentChar = Mid$(strCodeLine, lngCurrentIndex, 1)
                                Loop
                                If Len(strCurrentBuffer) > 0 Then    'its a compiler directive (we hope)
                                    'if the compiler directive starts with "#End" we want to flag
                                    'the entire line as a directive. you can also end in comment...
                                    If strCurrentBuffer = "End" Then
                                        lngBufferStart = lngCurrentIndex
                                        lngNextPlace = InStr(lngCurrentIndex + 1, strCodeLine, "If")
                                        If lngNextPlace Then
                                            lngCurrentIndex = lngBufferStart
                                            .Add MakeToken("#" & strCurrentBuffer & Mid$(strCodeLine, lngBufferStart, (lngNextPlace + 2) - lngCurrentIndex), TT_DIRECTIVE)
                                            lngCurrentIndex = lngNextPlace + 2
                                        Else
                                            .Add MakeToken("#" & strCurrentBuffer, TT_DIRECTIVE)
                                        End If

                                    Else
                                        .Add MakeToken("#" & strCurrentBuffer, TT_DIRECTIVE)
                                        'the fun part :)
                                        'if the compiler directive ends in "Then" we want to flag that to,
                                        'so we will recurse this function with everything in between starting
                                        'directive and "Then"  ...
                                        lngNextPlace = InStrRev(strCodeLine, "Then", , vbTextCompare)
                                        If lngNextPlace > lngCurrentIndex And lngNextPlace > InStrRev(strCodeLine, vbQuote) Then    'make sure its not in a quote
                                            Set RecurseList = Tokenize(Mid$(strCodeLine, lngCurrentIndex, lngNextPlace - lngCurrentIndex), lngLineNum)
                                            'check results
                                            If ParseError Then
                                                Exit Function
                                            End If
                                            If RecurseList.Count > 0 Then    'success!
                                                For Each Token In RecurseList    'add to token list
                                                    .Add Token
                                                Next
                                            End If
                                            .Add MakeToken("Then", TT_DIRECTIVE)
                                            lngCurrentIndex = lngNextPlace + 4
                                        End If
                                    End If
                                    'there is no "Then" or "If", parse as normal (probably a #Else)

                                Else    'not a compiler directive
                                    ParseError = True
                                    ErrorMessage = "A syntax or parse error has occurred on line #" & CStr(lngLineNum) & ". An invalid operator ( # ) was found with incorrect parameters. Was a compiler directive intended? Please check your code and try again." & vbCrLf & vbCrLf & "[ debug: " & Left$(strCodeLine, Len(strCodeLine) - 2) & " ]"
                                    Exit Function
                                End If
                            End If
                        Else
                            If strCurrentChar = "$" And NextChar(lngCurrentIndex, strCodeLine) = vbQuote Then
                                'in some attributes you have stuff like $"form1.frx:0000"
                                .Add MakeToken("$", TT_VARTYPE)
                                lngCurrentIndex = lngCurrentIndex + 1
                            Else
                                ParseError = True
                                ErrorMessage = "A syntax or parse error has occurred on line #" & CStr(lngLineNum) & ". An invalid operator ( " & strCurrentChar & " ) was found with incorrect spacing. Please check your code and try again." & vbCrLf & vbCrLf & "[ debug: " & Left$(strCodeLine, Len(strCodeLine) - 2) & " ]"
                                Exit Function
                            End If    '/strCurrentChar = "$" /
                        End If    '/strCurrentChar = "#"
                    End If    '/Len(strCurrentBuffer) > 0 /
                    strCurrentBuffer = ""
                    lngBufferStart = lngCurrentIndex


                Case "*", "/", "\", "^", "+", "="    'single operators
                    If (Len(strCurrentBuffer) > 0) And (strCurrentChar <> "=") Then    'syntax error
                        'we allow the "=" sign here because some hidden properties
                        'are formatted weird, like BackColor= 434345
                        ParseError = True
                        ErrorMessage = "A syntax or parse error has occurred on line #" & CStr(lngLineNum) & ". An invalid operator ( " & strCurrentChar & " ) was found with incorrect spacing. Please check your code and try again." & vbCrLf & vbCrLf & "[ debug: " & Left$(strCodeLine, Len(strCodeLine) - 2) & " ]"
                        Exit Function
                    Else
                        If (Len(strCurrentBuffer) > 0) Then    'its a stoogie one
                            .Add ResolveToken(strCurrentBuffer)
                        End If
                        strCurrentBuffer = ""

                        strNextChar = NextChar(lngCurrentIndex, strCodeLine)
                        strPrevChar = PrevChar(lngCurrentIndex, strCodeLine)
                        If (strNextChar = " " Or IsAlpha(strNextChar) Or SafeNumberCheck(strNextChar)) And (strPrevChar = " " Or strPrevChar = vbTab Or strPrevChar = "" Or IsAlpha(strPrevChar) Or SafeNumberCheck(strPrevChar) Or strPrevChar = ":") Then
                            .Add MakeToken(strCurrentChar, TT_OPERATOR)
                            If strNextChar = " " Then
                                lngCurrentIndex = lngCurrentIndex + 1
                                .Add MakeToken(" ", TT_SPACE)
                            End If
                        Else    'syntax error
                            ParseError = True
                            ErrorMessage = "A syntax or parse error has occurred on line #" & CStr(lngLineNum) & ". An invalid operator ( " & strCurrentChar & " ) was found with incorrect spacing. Please check your code and try again." & vbCrLf & vbCrLf & "[ debug: " & Left$(strCodeLine, Len(strCodeLine) - 2) & " ]"
                            Exit Function
                        End If
                    End If
                    lngCurrentIndex = lngCurrentIndex + 1
                    lngBufferStart = lngCurrentIndex


                Case "&"    'dual concat operator and type declarator
                    '&H00  &O11  varX & varY   991&  (991&)  (varX&) [varX&]? (varX&,varY&) ... (varX)& ?
                    'todo: rework and optimize logic
                    strNextChar = NextChar(lngCurrentIndex, strCodeLine)
                    strPrevChar = PrevChar(lngCurrentIndex, strCodeLine)
                    If strNextChar = " " And (strPrevChar = "" Or strPrevChar = " " Or strPrevChar = vbTab) Then    'concat operator
                        If Len(strCurrentBuffer) > 0 Then    'syntax/parse error
                            ParseError = True
                            ErrorMessage = "A syntax or parse error has occurred on line #" & CStr(lngLineNum) & ". An invalid operator ( " & strCurrentChar & " ) was found with incorrect spacing. Please check your code and try again." & vbCrLf & vbCrLf & "[ debug: " & Left$(strCodeLine, Len(strCodeLine) - 2) & " ]"
                            Exit Function
                        End If
                        .Add MakeToken(strCurrentChar, TT_OPERATOR)
                        If strNextChar = " " Then
                            lngCurrentIndex = lngCurrentIndex + 1
                            .Add MakeToken(" ", TT_SPACE)
                        End If

                    ElseIf (strNextChar = "H" Or strNextChar = "O") And (strPrevChar = " " Or strPrevChar = "(" Or strPrevChar = vbTab Or strPrevChar = "" Or strPrevChar = "[") Then    'hex or octal start
                        If Len(strCurrentBuffer) > 0 Then    'syntax error
                            ParseError = True
                            ErrorMessage = "A syntax or parse error has occurred on line #" & CStr(lngLineNum) & ". An invalid operator ( " & strCurrentChar & " ) was found with incorrect spacing. Please check your code and try again." & vbCrLf & vbCrLf & "[ debug: " & Left$(strCodeLine, Len(strCodeLine) - 2) & " ]"
                            Exit Function
                        End If
                        .Add MakeToken(strCurrentChar, TT_VARTYPE)
                    Else
                        If (SafeNumberCheck(strPrevChar) Or IsAlpha(strPrevChar)) And (strNextChar = " " Or strNextChar = ")" Or strNextChar = vbCr Or strNextChar = "," Or strNextChar = ":" Or strNextChar = "]" Or strNextChar = "(") Then    'type declaration
                            If Len(strCurrentBuffer) > 0 Then
                                .Add ResolveToken(strCurrentBuffer)
                                strCurrentBuffer = ""
                            End If
                            .Add MakeToken(strCurrentChar, TT_VARTYPE)
                            If strNextChar = " " Then
                                lngCurrentIndex = lngCurrentIndex + 1
                                .Add MakeToken(" ", TT_SPACE)
                            End If
                        Else
                            ParseError = True
                            ErrorMessage = "A syntax or parse error has occurred on line #" & CStr(lngLineNum) & ". An invalid operator ( " & strCurrentChar & " ) was found with incorrect spacing. Please check your code and try again." & vbCrLf & vbCrLf & "[ debug: " & Left$(strCodeLine, Len(strCodeLine) - 2) & " ]"
                            Exit Function
                        End If
                    End If
                    lngCurrentIndex = lngCurrentIndex + 1
                    lngBufferStart = lngCurrentIndex


                Case "-"    'dual subtract operator or negator
                    'todo: rework logic for this , currently kludged
                    strNextChar = NextChar(lngCurrentIndex, strCodeLine)
                    strPrevChar = PrevChar(lngCurrentIndex, strCodeLine)
                    If Len(strCurrentBuffer) > 0 Then
                        If ((SafeNumberCheck(strNextChar) Or IsAlpha(strNextChar)) And (SafeNumberCheck(strPrevChar) Or IsAlpha(strPrevChar))) Then
                            'this allows for things like: {248DD890-BB45-11CF-9ABC-0080C7E7B78D}
                            strCurrentBuffer = strCurrentBuffer & strCurrentChar
                            lngCurrentIndex = lngCurrentIndex + 1
                        Else    'syntax/parse error
                            ParseError = True
                            ErrorMessage = "A syntax or parse error has occurred on line #" & CStr(lngLineNum) & ". An invalid operator ( " & strCurrentChar & " ) was found with incorrect spacing. Please check your code and try again." & vbCrLf & vbCrLf & "[ debug: " & Left$(strCodeLine, Len(strCodeLine) - 2) & " ]"
                            Exit Function
                        End If
                    Else
                        If (strPrevChar = " " Or strPrevChar = "" Or strPrevChar = vbTab Or strPrevChar = "(" Or strPrevChar = ",") And strNextChar = " " Then    'operator
                            .Add MakeToken(strCurrentChar, TT_OPERATOR)
                            lngCurrentIndex = lngCurrentIndex + 1
                            .Add MakeToken(" ", TT_SPACE)
                        Else
                            If SafeNumberCheck(strNextChar) Then
                                .Add MakeToken(strCurrentChar, TT_NEGATIVE)
                            ElseIf (IsAlpha(strNextChar) Or strNextChar = "(" Or strNextChar = " ") And (SafeNumberCheck(strPrevChar) Or IsAlpha(strPrevChar) Or strPrevChar = "(" Or strPrevChar = " " Or strPrevChar = "[" Or strPrevChar = vbTab Or strPrevChar = ")") Then
                                'todo: re work logic for this , currently kludged
                                .Add MakeToken(strCurrentChar, TT_NEGATIVE)
                            Else    'syntax error
                                ParseError = True
                                ErrorMessage = "A syntax or parse error has occurred on line #" & CStr(lngLineNum) & ". An invalid operator ( " & strCurrentChar & " ) was found with incorrect spacing. Please check your code and try again." & vbCrLf & vbCrLf & "[ debug: " & Left$(strCodeLine, Len(strCodeLine) - 2) & " ]"
                                Exit Function
                            End If
                        End If
                        lngCurrentIndex = lngCurrentIndex + 1
                        lngBufferStart = lngCurrentIndex
                    End If


                Case "0"    'for: GoTo 0
                    'todo: add case for all GoTo <Number>'s
                    lngCurrentIndex = lngCurrentIndex + 1
                    If Len(strCurrentBuffer) = 0 Then
                        If InStr(strCodeLine, "GoTo 0") + 6 = lngCurrentIndex Then
                            .Add MakeToken("0", TT_KEYWORD)
                            strCurrentChar = ""
                            lngBufferStart = lngCurrentIndex
                        End If
                    End If
                    strCurrentBuffer = strCurrentBuffer & strCurrentChar


                Case "(", ",", ")", vbTab, vbCr    'all reduce/shift types
                    If Len(strCurrentBuffer) > 0 Then
                        .Add ResolveToken(strCurrentBuffer)
                        strCurrentBuffer = ""
                    End If

                    If strCurrentChar = vbCr Then    'end of line
                        .Add MakeToken(vbCrLf, TT_EOL)
                        Exit Function
                    End If

                    .Add ResolveToken(strCurrentChar)
                    lngCurrentIndex = lngCurrentIndex + 1
                    lngBufferStart = lngCurrentIndex


                Case ":", " "
                    If Len(strCurrentBuffer) > 0 Then
                        If (lngBufferStart = 1) And SafeNumberCheck(strCurrentBuffer) And (strCurrentChar = ":") Then
                            'unbelievably, someone is using line numbers !
                            .Add MakeToken(strCurrentBuffer, TT_LINENUM)
                        Else
                            'probably a goto <label>:  [ tsk... or bad coding style! :) ]
                            .Add ResolveToken(strCurrentBuffer)
                        End If
                        strCurrentBuffer = ""
                    End If

                    If strCurrentChar = " " Then
                        .Add MakeToken(strCurrentChar, TT_SPACE)
                    Else
                        .Add MakeToken(strCurrentChar, TT_LINESTART)
                    End If
                    lngCurrentIndex = lngCurrentIndex + 1
                    lngBufferStart = lngCurrentIndex


                Case "[", "]"    'forced names e.g   mCol.[_NewEnum]
                    If Len(strCurrentBuffer) > 0 Then
                        .Add ResolveToken(strCurrentBuffer)
                        strCurrentBuffer = ""
                    End If

                    If strCurrentChar = "[" Then
                        .Add MakeToken(strCurrentChar, TT_FORCEVARBEGIN)
                    Else
                        .Add MakeToken(strCurrentChar, TT_FORCEVAREND)
                    End If

                    lngCurrentIndex = lngCurrentIndex + 1
                    lngBufferStart = lngCurrentIndex


                Case "_"    'line continue or hidden interface "mCol.[_NewEnum]" or identifier as in TOKEN_TYPE
                    If NextChar(lngCurrentIndex, strCodeLine) = vbCr Then    'eol: its probably a line continue
                        If Len(strCurrentBuffer) > 0 Then
                            .Add ResolveToken(strCurrentBuffer)
                            strCurrentBuffer = ""
                        End If
                        'we are done
                        .Add MakeToken(strCurrentChar, TT_LINECONTINUE)
                        .Add MakeToken(vbCrLf, TT_EOL)
                        Exit Function
                    Else
                        strCurrentBuffer = strCurrentBuffer & strCurrentChar
                        lngCurrentIndex = lngCurrentIndex + 1
                    End If


                Case vbQuote    'beginning a string literal/constant
                    If Len(strCurrentBuffer) > 0 Then
                        .Add ResolveToken(strCurrentBuffer)
                    End If
                    .Add MakeToken(vbQuote, TT_STRINGBEGIN)
                    strCurrentBuffer = ""

                    lngCurrentIndex = lngCurrentIndex + 1
                    lngBufferStart = lngCurrentIndex
                    'loop thru rest of string
                    strCurrentChar = Mid$(strCodeLine, lngCurrentIndex, 1)
                    Do While strCurrentChar <> vbQuote And lngCurrentIndex <= lngCodeLength
                        strCurrentBuffer = strCurrentBuffer & strCurrentChar
                        lngCurrentIndex = lngCurrentIndex + 1
                        strCurrentChar = Mid$(strCodeLine, lngCurrentIndex, 1)
                    Loop
                    'any text?
                    If Len(strCurrentBuffer) > 0 Then
                        .Add MakeToken(strCurrentBuffer, TT_STRINGTEXT)
                        strCurrentBuffer = ""
                        lngBufferStart = lngCurrentIndex
                    End If
                    'current char should be another end quote
                    If strCurrentChar = vbQuote Then
                        .Add MakeToken(vbQuote, TT_STRINGEND)
                        strCurrentBuffer = ""
                        lngCurrentIndex = lngCurrentIndex + 1
                        lngBufferStart = lngCurrentIndex
                    Else    'shouldnt happen, means no ending quote :(
                        ParseError = True
                        ErrorMessage = "A syntax error has occurred on line #" & CStr(lngLineNum) & ". No matching end of string found. Please check your code and try again." & vbCrLf & vbCrLf & "[ debug: " & Left$(strCodeLine, Len(strCodeLine) - 2) & " ]"
                        Exit Function
                    End If


                Case "'"    'a comment, nothing else in this line ...(except possibly line continue!)
                    If Len(strCurrentBuffer) > 0 Then    'must be a syntax/parse error...
                        ParseError = True
                        ErrorMessage = "A syntax error has occurred on line #" & CStr(lngLineNum) & ". Was a comment intended?. Please check your code and try again." & vbCrLf & vbCrLf & "[ debug: " & Left$(strCodeLine, Len(strCodeLine) - 2) & " ]"
                        Exit Function
                    End If
                    If BlnNextLineIsComment Then
                        BlnNextLineIsComment = False
                    Else
                        .Add MakeToken("'", TT_COMMENTSTART)
                    End If
                    'move pointer past -->'
                    lngCurrentIndex = lngCurrentIndex + 1

                    If lngCurrentIndex <= lngCodeLength - 2 Then    'there is comment
                        'look for line continue
                        lngNextPlace = InStrRev(strCodeLine, " _" & vbCrLf)
                        If lngNextPlace > lngCurrentIndex Then
                            lngNextPlace = lngNextPlace + 1
                            BlnNextLineIsComment = True    'set flag for next line of code
                        Else
                            lngNextPlace = lngCodeLength - 1
                        End If
                        If lngNextPlace > lngCurrentIndex Then
                            .Add MakeToken(Mid$(strCodeLine, lngCurrentIndex, lngNextPlace - lngCurrentIndex), TT_COMMENTTEXT)
                        End If
                        If BlnNextLineIsComment Then
                            .Add MakeToken("_", TT_LINECONTINUE)
                        End If
                    End If

                    'finished
                    .Add MakeToken(vbCrLf, TT_EOL)
                    Exit Function


                Case "R"
                    Rem - a PITA comment! _
                     these are evaluated a step at a time, to avoid unnesacary comparisons
                    If (NextChar(lngCurrentIndex, strCodeLine) = "e") Then
                        If (NextChar(lngCurrentIndex + 1, strCodeLine) = "m") Then
                            If (NextChar(lngCurrentIndex + 2, strCodeLine) = " ") Then

                                If Len(strCurrentBuffer) > 0 Then    'syntax error
                                    ParseError = True
                                    ErrorMessage = "A syntax error has occurred on line #" & CStr(lngLineNum) & ". Was a comment intended?. Please check your code and try again." & vbCrLf & vbCrLf & "[ debug: " & Left$(strCodeLine, Len(strCodeLine) - 2) & " ]"
                                    Exit Function
                                End If

                                If BlnNextLineIsComment Then
                                    BlnNextLineIsComment = False
                                Else
                                    .Add MakeToken("Rem", TT_COMMENTSTART)
                                End If

                                'move pointer to end of "Rem "
                                lngCurrentIndex = lngCurrentIndex + 3

                                If lngCurrentIndex <= lngCodeLength - 2 Then    'there is comment
                                    'look for line continue
                                    lngNextPlace = InStrRev(strCodeLine, " _" & vbCrLf)
                                    If lngNextPlace Then
                                        lngNextPlace = lngNextPlace + 1
                                        BlnNextLineIsComment = True    'set flag for next line of code
                                    Else
                                        lngNextPlace = lngCodeLength - 1
                                    End If
                                    If lngNextPlace > lngCurrentIndex Then
                                        .Add MakeToken(Mid$(strCodeLine, lngCurrentIndex, lngNextPlace - lngCurrentIndex), TT_COMMENTTEXT)
                                    End If
                                End If

                                'we are done
                                .Add MakeToken(vbCrLf, TT_EOL)
                                Exit Function

                            End If    '/(NextChar(lngCurrentIndex + 2, strCodeLine) = " ")
                        End If    '/(NextChar(lngCurrentIndex + 1, strCodeLine) = "m")/
                    End If    '/(NextChar(lngCurrentIndex, strCodeLine) = "e")/

                    strCurrentBuffer = strCurrentBuffer & strCurrentChar
                    lngCurrentIndex = lngCurrentIndex + 1


                Case Else
                    strCurrentBuffer = strCurrentBuffer & strCurrentChar
                    If Len(strCurrentBuffer) > MAX_LEN Then    'we probably have a problem (default: 50)
                        'myverylongobjectmightbecalling.myverylongfunction()
                        ' we should never get here
                        If MsgBox("This line of code: [ " & strCurrentBuffer & " ] at line " & CStr(lngLineNum) & ", is unable to be determined by VTH. It could be a syntax error or not implemented by VTH. If you continue, this line of code will possibly not be syntax highlighted. Continue?", vbCritical + vbYesNo) = vbYes Then
                            'reduce and continue
                            .Add ResolveToken(strCurrentBuffer)    'hope for the best :)
                            strCurrentBuffer = ""
                            'find a new safe starting point, uhuh...
                            lngCurrentIndex = FindStartPlace(lngCurrentIndex, strCodeLine, lngCodeLength)
                            lngBufferStart = lngCurrentIndex
                        Else
                            'error out
                            ParseError = True
                            ErrorMessage = "Conversion operation aborted... Unable to resolve code."
                            Exit Function
                        End If
                    Else
                        lngCurrentIndex = lngCurrentIndex + 1
                    End If

            End Select
        Loop

        'just in case we missed something
        If Len(strCurrentBuffer) > 0 Then
            .Add ResolveToken(strCurrentBuffer)
        End If

    End With '/Tokenize/


End Function

'this function (used by the Tokenize() function) looks at a token and decides
'what kind of token it is... it helps keep the logic simpler in Tokenize()
Function ResolveToken(ByVal strToken As String) As TokenItem
    'Written by Lewis Miller

    Dim lngFindPlace    As Long
    Dim lngTokenLength  As Long
    Dim strFirstChar    As String
    Dim strSecondChar   As String

    'we are here because we need to resolve what Tokenize() couldnt identify with normal constants
    'we need to reduce the following   (  ) ,  vbTab  vbCr  : space    and any keywords/variables

    lngTokenLength = Len(strToken)
    If lngTokenLength = 0 Then    'fail
        Set ResolveToken = New TokenItem    'empty token
    Else
        Select Case strToken
            Case "("
                Set ResolveToken = MakeToken(strToken, TT_LISTBEGIN)
                Exit Function
            Case ")"
                Set ResolveToken = MakeToken(strToken, TT_LISTEND)
               Exit Function
            Case ","
                Set ResolveToken = MakeToken(strToken, TT_LISTSEPERATOR)
                Exit Function
            Case vbTab
                Set ResolveToken = MakeToken(strToken, TT_TAB)
                Exit Function
            Case vbCr
                Set ResolveToken = MakeToken(vbCrLf, TT_EOL)
                Exit Function
                
            Case ":"
                Set ResolveToken = MakeToken(strToken, TT_LINESTART)
                Exit Function
                
            Case " "
                Set ResolveToken = MakeToken(strToken, TT_SPACE)
                Exit Function
                           
                'all easy ones done
            Case Else
                If isKeyword(strToken) Then    'is it a keyword?
                    Set ResolveToken = MakeToken(strToken, TT_KEYWORD)
                    Exit Function
                ElseIf SafeNumberCheck(strToken) Then    'is it numeric?
                    Set ResolveToken = MakeToken(strToken, TT_NUMERIC)
                    Exit Function
                Else
                    If lngTokenLength = 1 Then    'only one char

                        'we could be dealing with a weird eol - chr(10)
                        If strToken = vbLf Then
                            'convert it
                            Set ResolveToken = MakeToken(vbCrLf, TT_EOL)
                           Exit Function
                        End If

                        If IsAlpha(strToken) Then    'single letter variable?
                            Set ResolveToken = MakeToken(strToken, TT_VARIABLE)
                            Exit Function
                        End If

                        'todo: more decoding for single char tokens here?

                    Else    '//multi character token

                        strFirstChar = Left$(strToken, 1)    'first char
                        strSecondChar = Mid$(strToken, 2)    'second char


                        If strFirstChar = "H" Or strFirstChar = "O" Then    'probably a hex or octal constant
                            If SafeNumberCheck(strSecondChar) Or (strSecondChar Like "[A-F,a-f]") Then
                                Set ResolveToken = MakeToken(strToken, TT_HEXOROCTAL)
                               Exit Function
                            End If
                        End If

                        If strFirstChar = "." And IsAlpha(strSecondChar) Then    'we are inside a With statement
                            Set ResolveToken = MakeToken(strToken, TT_SCOPEDCALL)
                            Exit Function
                        End If

                        lngFindPlace = InStr(strToken, ".")    'scoped call?  e.g  modIni.ReadNumber
                        If (lngFindPlace > 1) And (lngFindPlace < lngTokenLength) And IsAlpha(strFirstChar) Then
                            Set ResolveToken = MakeToken(strToken, TT_SCOPEDCALL)
                            Exit Function
                        End If

                        'anything else is treated as a variable for now
                        Set ResolveToken = MakeToken(strToken, TT_VARIABLE)
                        Exit Function
                        
                    End If
                End If

        End Select

        If ResolveToken Is Nothing Then    'nothing found
            'return default/unknown for now...
            Set ResolveToken = MakeToken(strToken, TT_UNKNOWN)
        End If

    End If

End Function

'this is a simple helper function to create a token object...
'it helps keep the logic simpler in the Tokenize() function
Function MakeToken(ByVal strToken As String, enmTokenKind As TOKEN_TYPE) As TokenItem
    'Written by Lewis Miller

    Set MakeToken = New TokenItem

    With MakeToken
        .Kind = enmTokenKind
        .Value = strToken
    End With

End Function

'safely checks a string to see if it is numeric
Function SafeNumberCheck(ByVal strNumber As String) As Boolean
    'Written by Lewis Miller

    Dim lngNumLength As Long

    lngNumLength = Len(strNumber)
    
    If (lngNumLength > 0) And (lngNumLength < 17) Then    'supports up to a 16 byte (128 bit) number
        SafeNumberCheck = IsNumeric(strNumber)
    End If

End Function

'this functions sole purpose is to find a safe new starting point
'after a parse error in a line of code. *very rarely* will this be called, if ever.
Function FindStartPlace(ByVal lngStartingPoint As Long, ByVal strCodeLine As String, ByVal lngCodeLen As Long) As Long
    'Written by Lewis Miller

    Dim strCurrentChar  As String
    Dim blnInQuote      As Boolean
    Dim lngLoopIndex    As Long

    If lngCodeLen < 2 Then
        GoTo SkipAllThat
    End If

    If lngStartingPoint > lngCodeLen - 1 Then
        FindStartPlace = lngCodeLen - 1
        Exit Function
    End If

    For lngLoopIndex = 1 To lngCodeLen
        strCurrentChar = Mid$(strCodeLine, lngLoopIndex, 1)
        Select Case strCurrentChar
            Case """"
                blnInQuote = Not blnInQuote
                If (lngLoopIndex > lngStartingPoint) And (Not blnInQuote) Then
                    FindStartPlace = lngLoopIndex + 1
                    Exit Function
                End If

                'safe starting points
            Case " ", vbTab, vbCr, "&", "(", ")", ",", "-", "*", "^", "+", "=", "<", ">", "/", "\", "[", "]", ":"
                If lngLoopIndex > lngStartingPoint And (Not blnInQuote) Then
                    FindStartPlace = lngLoopIndex
                    Exit Function
                End If
        End Select
    Next lngLoopIndex

SkipAllThat:

    'nothing satisfactory found, return start place
    If lngStartingPoint + 1 <= lngCodeLen Then
        FindStartPlace = lngStartingPoint + 1
    Else
        FindStartPlace = lngStartingPoint
    End If

End Function

'test a character to see if it is an english alphabet letter
Function IsAlpha(ByVal strTest As String, Optional ByVal TestUpper As Boolean, Optional ByVal TestLower As Boolean) As Boolean
    'Written by Lewis Miller

    If Len(strTest) > 0 Then
        If TestUpper Then
            IsAlpha = (strTest Like "[A-Z]")    'upper case
        ElseIf TestLower Then
            IsAlpha = (strTest Like "[a-z]")    'lower case
        Else
            IsAlpha = (strTest Like "[A-Z,a-z]")    'normal
        End If
    End If

End Function

'determines if a single character is the starting character of a line of code after whitespace
Function IsFirstChar(ByVal strFindChar As String, ByVal strCodeLine As String, ByVal lngCodeLen As Long) As Boolean
    'Written by Lewis Miller

    Dim lngCharIndex As Long
    Dim strCurChar   As String

    lngCharIndex = 1
    strCurChar = Mid$(strCodeLine, lngCharIndex, 1)
    Do While (strCurChar = " " Or strCurChar = vbTab) And lngCharIndex < lngCodeLen
        lngCharIndex = lngCharIndex + 1
        strCurChar = Mid$(strCodeLine, lngCharIndex, 1)
    Loop
    IsFirstChar = (strCurChar = strFindChar)

End Function

'determines if a string is the starting token of a line of code, after whitespace...
Function IsFirstToken(ByVal strFind As String, ByVal strCodeLine As String, ByVal lngCodeLen As Long) As Boolean
    'Written by Lewis Miller

    Dim lngCharIndex As Long
    Dim strCurChar   As String
    Dim lngFindLen   As Long

    lngFindLen = Len(strFind)
    If lngFindLen > 0 Then

        lngCharIndex = 1
        strCurChar = Mid$(strCodeLine, lngCharIndex, 1)
        Do While (strCurChar = " " Or strCurChar = vbTab) And (lngCharIndex < lngCodeLen - (lngFindLen + 1))
            lngCharIndex = lngCharIndex + 1
            strCurChar = Mid$(strCodeLine, lngCharIndex, 1)
        Loop

        IsFirstToken = (strFind = Mid$(strCodeLine, lngCharIndex, lngFindLen))
    End If

End Function

'determines if a string is the ending token of a line of code, before whitespace...
'strFind cannot end in whitespace i.e.  vbcr vblf, space, tab
Function IsLastToken(ByVal strFind As String, ByVal strCodeLine As String, ByVal lngCodeLen As Long) As Boolean
    'Written by Lewis Miller

    Dim lngCharIndex As Long
    Dim strCurChar   As String
    Dim lngFindLen   As Long

    lngFindLen = Len(strFind)
    If lngFindLen > 0 Then
        lngCharIndex = lngCodeLen

        strCurChar = Mid$(strCodeLine, lngCharIndex, 1)
        Do While (strCurChar = " " Or strCurChar = vbTab Or strCurChar = vbLf Or strCurChar = vbCr) And (lngCharIndex > 0 Or lngCharIndex > lngCodeLen - (lngFindLen + 1))
            lngCharIndex = lngCharIndex - 1
            strCurChar = Mid$(strCodeLine, lngCharIndex, 1)
        Loop
        If lngCharIndex >= lngFindLen Then
           IsLastToken = (strFind = Mid$(strCodeLine, (lngCharIndex - lngFindLen) + 1, lngFindLen))
        End If
    End If

End Function


'this function is the main controller of most of the functions in this module
Function ConvertCode(ByVal strInput As String) As Boolean
    'Written by Lewis Miller

    Dim Tokens          As TokenList    'list of tokens
    Dim CodeArr()       As String    'an array of lines of code
    Dim lngLoopIndex    As Long    'loop counter
    Dim lngArrBound     As Long    'upper bound of CodeArr()

    'the following variables are all used for determining if
    'a line of code is the beginning/end of a sub,function, or property
    'in order to add a horizontal line if that option is checked
    Dim blnFuncSeen     As Boolean
    Dim lngEndStart     As Long
    Dim lngFuncStart    As Long
    Dim lngFuncLen      As Long
    Dim lngListStart    As Long
    Dim lngListEnd      As Long
    Dim lngQuoteStart   As Long
    Dim lngCommaStart   As Long
    Dim lngCommentStart As Long
    
    'these variables are used for looking back and forward in the
    'codearr() for an insertion point to insert the horz line
    Dim lngInsertIndex  As Long
    Dim lngSearchIndex  As Long
    Dim strCodeLine     As String
    Dim lngCodeLen      As Long

    With frmMain     'this is why having a public main form is a good thing....

        'make sure there is code
        If Len(strInput) > 0 Then

            'store the current time
            StartTime = timeGetTime

            'prepare list for lines of html (table start, etc)
            PrepStack

            'set status text
            If Not blnCmdMode Then
                .lblOverallStatus = "Separating code into lines..."
            End If

            'check to see if there is more than one line of code
            If InStr(strInput, vbCrLf) Then
                CodeArr = Split(strInput, vbCrLf)
            Else
                'only one line of code
                ReDim CodeArr(0) As String
                CodeArr(0) = strInput
            End If

            'get and store upper bound of array
            lngArrBound = UBound(CodeArr)
            'reset progress bar
            If Not blnCmdMode Then
                .SetOverallProgressMax lngArrBound + 1
            End If

            'loop through each line of code and convert
            For lngLoopIndex = 0 To lngArrBound
                'set status text
                If Not blnCmdMode Then
                    .lblOverallStatus = "Converting Line #" & CStr(lngLoopIndex + 1) & " of " & CStr(lngArrBound + 1)
                End If

                're-add the end of line
                CodeArr(lngLoopIndex) = CodeArr(lngLoopIndex) & vbCrLf
                
                strCodeLine = CodeArr(lngLoopIndex)
                lngCodeLen = Len(CodeArr(lngLoopIndex))
                
                'all this, to see if we have to add a horizontal line...
                'examine the code line to see if its the beginning of a sub,function, or property
                'we have to do alot of checking, since we are 'string searching' the line
                'and we have to make sure its what we want...
                If blnAddLine And (Len(strCodeLine) > 6) And (Not BlnNextLineIsComment) Then
                    If Not blnFuncSeen Then
                        'this is the first sub or function, we will want to add a preceding line
                        'only if there is general dec's, this should happen only once per file

                        '...find our flag
                        lngFuncStart = InStr(1, strCodeLine, "Sub ")
                        lngFuncLen = 3
                        If lngFuncStart = 0 Then
                            lngFuncStart = InStr(1, strCodeLine, "Function ")
                            lngFuncLen = 8
                            If lngFuncStart = 0 Then
                                lngFuncStart = InStr(1, strCodeLine, "Property ")
                            End If
                        End If

                        'did we find anything?
                        If (lngFuncStart > 0) And (InStr(strCodeLine, "Declare ") = 0) Then    'not an API call
                            'all function , sub, props have these
                            lngListStart = InStr(strCodeLine, "(")
                            If (lngListStart > 0) And (lngListStart > lngFuncStart + lngFuncLen + 1) Then
                                lngListEnd = InStr(strCodeLine, ")")
                                If (lngListEnd > 0) And (lngListEnd > lngListStart) Then
                                    'look for quoted text
                                    lngQuoteStart = InStr(strCodeLine, vbQuote)
                                    If (lngQuoteStart > lngListStart) Or (lngQuoteStart = 0) Then
                                        'look for a comment
                                        lngCommentStart = InStr(strCodeLine, "'")
                                        If (lngCommentStart = 0) Or (lngCommentStart > lngListEnd) Then
                                        lngCommentStart = InStr(strCodeLine, " Rem ")
                                        If (lngCommentStart = 0) Or (lngCommentStart > lngListEnd) Then
                                            lngCommaStart = InStr(strCodeLine, ",")
                                              If (lngCommaStart = 0) Or ((lngCommaStart > lngListStart) And (lngCommaStart > (lngFuncStart + lngFuncLen + 1))) Then
                                                'well we are pretty sure now...
                                                lngInsertIndex = Html.ListCount    'last line 'pushed' + 1
                                                'now we have to find the end of the general declarations section,
                                                'or a point where there are no more comments or empty lines, backwards...
                                                lngSearchIndex = lngLoopIndex - 1    'back one line
                                                If lngLoopIndex > 0 And lngSearchIndex >= 0 And lngInsertIndex > 0 Then
                                                    'loop backwards
                                                    Do While lngInsertIndex >= 0 And lngSearchIndex >= 0
                                                        strCodeLine = CodeArr(lngSearchIndex)
                                                        lngCodeLen = Len(CodeArr(lngSearchIndex))
                                                        If lngCodeLen > 2 Then
                                                            If IsFirstChar("'", strCodeLine, lngCodeLen) Or IsFirstToken("Rem ", strCodeLine, lngCodeLen) Then
                                                                lngInsertIndex = lngInsertIndex - 1
                                                            Else
                                                                Exit Do
                                                            End If
                                                        Else
                                                            If strCodeLine = vbCrLf Then
                                                                lngInsertIndex = lngInsertIndex - 1
                                                            Else
                                                                Exit Do
                                                            End If
                                                        End If
                                                        lngSearchIndex = lngSearchIndex - 1
                                                    Loop
                                                End If
                                                If lngInsertIndex > 0 And lngSearchIndex >= 0 Then    'we have found an acceptable spot
                                                    'just like a listbox, we add(insert) a horz line to the string list with an index
                                                    Html.AddItem vbCrLf & "<hr size=" & vbQuote & "2" & vbQuote & " align=" & vbQuote & "left" & vbQuote & "><br>" & vbCrLf, lngInsertIndex
                                                End If
                                                strCodeLine = CodeArr(lngLoopIndex)
                                                lngCodeLen = Len(strCodeLine)
                                                blnFuncSeen = True
                                            End If '/lngCommaStart = 0/
                                        End If '/lngCommentStart = 0
                                        End If '/lngCommentStart = 0
                                    End If '/lngQuoteStart > lngListStart/
                                End If '/lngListEnd > 0/
                            End If '/lngListStart > 0/
                        End If '/lngFuncStart > 0/
                    End If '/Not blnFuncSeen/
                End If '/blnAddLine/
                
                'thats alot of ifs...

                'call the tokenizer engine to tokenize the line of code into a list of tokens
                Set Tokens = Tokenize(strCodeLine, lngLoopIndex + 1)

                'has a parse error occured?
                If ParseError Then
                    MsgBox ErrorMessage, vbCritical
                    If Not blnCmdMode Then 'not command line mode
                        'set status on frmMain
                        .lblOverallStatus = "Parse Error On  Line #" & CStr(lngLoopIndex + 1)
                    End If
                    GoTo Done
                Else
                    'now see if we want to close off with a succeeding horizontal line
                    'we look for End Sub, End Function, or End Property.
                    If blnAddLine And (lngCodeLen > 6) And (TotalLines > 0) And (lngLoopIndex < lngArrBound - 1) Then
                        lngEndStart = InStr(strCodeLine, "End ")
                        'this should only happen after a line has already been added to general dec's
                        If blnFuncSeen And (lngEndStart > 0) Then
                            lngFuncStart = InStr(1, strCodeLine, " Sub")
                            lngFuncLen = 3
                            If lngFuncStart = 0 Then
                                lngFuncStart = InStr(1, strCodeLine, " Function")
                                lngFuncLen = 8
                                If lngFuncStart = 0 Then
                                    lngFuncStart = InStr(1, strCodeLine, " Property")
                                End If
                            End If
                            If (lngFuncStart > 0) Then
                                If lngFuncStart = lngEndStart + 3 Then
                                    lngCommentStart = InStr(strCodeLine, "'")
                                    If (lngCommentStart = 0) Or (lngCommentStart > lngFuncStart + lngFuncLen + 1) Then
                                    lngCommentStart = InStr(strCodeLine, "Rem ")
                                    If (lngCommentStart = 0) Or (lngCommentStart > lngFuncStart + lngFuncLen + 1) Then
                                        lngQuoteStart = InStr(strCodeLine, vbQuote)
                                        If (lngQuoteStart > lngCommentStart) Or (lngQuoteStart = 0) Then
                                            'now we want to make sure this isnt the last sub or function
                                            'because we dont want to add a line after those...
                                            lngSearchIndex = lngLoopIndex + 1
                                            lngInsertIndex = 0
                                            'loop through the rest of code array looking for end func constructs
                                            Do While lngSearchIndex < lngArrBound And lngInsertIndex = 0
                                                strCodeLine = CodeArr(lngSearchIndex)
                                                lngCodeLen = Len(strCodeLine)
                                                
                                                'we look for End Sub, End Function, End Property starting with most likely one
                                                If IsFirstToken("End Sub", strCodeLine, lngCodeLen) Then
                                                    lngInsertIndex = 1    'flag (re-using a variable)
                                                    Exit Do 'abort loop, no need to look for the others
                                                End If
                                                If IsFirstToken("End Function", strCodeLine, lngCodeLen) Then
                                                    lngInsertIndex = 1    'flag (re-using a variable)
                                                    Exit Do
                                                End If
                                                If IsFirstToken("End Property", strCodeLine, lngCodeLen) Then
                                                    lngInsertIndex = 1    'flag (re-using a variable)
                                                    Exit Do
                                                End If
                                                lngSearchIndex = lngSearchIndex + 1
                                            Loop
                                            If lngInsertIndex > 0 Then    'an "End <Func>" was found
                                            'note: tokens cannot be 0 length, so a dash is used to fill in
                                                Tokens.Add MakeToken("-", TT_DIVIDER)
                                            End If

                                        End If '/lngQuoteStart > lngCommentStart/
                                    End If '/lngCommentStart = 0/
                                    End If '/lngCommentStart = 0/
                                End If '/lngFuncStart = lngEndStart + 3/
                            End If '/lngFuncStart > 0/
                        End If '/blnFuncSeen/
                    End If '/blnAddLine
                    
                    'increment global line count
                    TotalLines = TotalLines + 1    'this assumes all callers reset to 0
                    'convert the tokens into html and push onto the html list
                    Html.AddItem ConvertTokenListToHtml(Tokens)
                End If

                If (lngLoopIndex + 1) Mod 15 = 0 Then
                    If Not blnCmdMode Then
                        .SetOverallProgress lngLoopIndex + 1
                    End If
                    If blnCancelled Then Exit For 'check for cancellation
                    
                End If
                
                'keep from freezing on large files
                If (lngLoopIndex Mod 250) = 0 Then
                    DoEvents
                End If
                
            Next lngLoopIndex

            If Not blnCmdMode Then
                .lblOverallStatus = "Converted " & CStr(TotalLines) & " Lines In " & CalculateTime(timeGetTime - StartTime)
            End If

            ConvertCode = True
Done:
            'close off list with ending html ( </table></html> etc... )
            CloseStack
        End If

        If Not blnCmdMode Then
            .SetOverallProgressMax 0
        End If
    End With

End Function

'this function loads a file buffer and calls convertcode() on it
Function ConvertFile(ByVal strInputFile As String, ByVal strOutputFile As String, ByVal blnOnlySource As Boolean, ByVal blnAddPreview As Boolean, Optional blnOverwrite As Boolean) As Boolean

    Dim strInput        As String
    Dim lngPosition     As Long
    Dim intFileNum      As Integer
    Dim lngLoopIndex    As Long
    Dim lngListCount    As Long

    If (Len(strInputFile)) = 0 Or (Not FileCanRead(strInputFile)) Then
        MsgBox "Invalid input file (" & GetFileName(strInputFile) & "). Cannot read from, or doesnt exist. Please check the file path and try again.", vbCritical
        Exit Function
    End If
    If (Len(strOutputFile)) = 0 Then
        MsgBox "Invalid output file (" & GetFileName(strOutputFile) & "). Please check the file path and try again.", vbCritical
        Exit Function
    End If
    If FileCanRead(strOutputFile) Then
        If Not blnOverwrite Then    'prompt user for validation to delete
            If MsgBox("Output file (" & GetFileName(strOutputFile) & ") already exists. Do you wish to continue and overwrite any data in the file?.", vbCritical + vbYesNo) = vbNo Then
                Exit Function
            End If
        End If
    End If

    strInput = FileToString(strInputFile)

    If Len(strInput) > 0 Then
        'most all vb files have this: "Attribute VB_"
        lngPosition = InStr(1, strInput, "Attribute VB_")
        If lngPosition = 0 Then    'not found
            If MsgBox(GetFileName(strInputFile) & " does not appear to be a valid visual basic 6 source code file! Do you want to continue?", vbCritical + vbYesNo) = vbNo Then
                Exit Function
            End If
        End If

        If blnOnlySource Then    'remove VB file attributes and properties (this is actually faster if they are removed, then processing the attributes)
            If lngPosition Then
                lngPosition = InStr(lngPosition + 1, strInput, vbCrLf)
                strInput = Mid$(strInput, lngPosition + 2)
                Do While Left$(strInput, 13) = "Attribute VB_"
                    lngPosition = InStr(1, strInput, vbCrLf, vbBinaryCompare)
                    If lngPosition > 0 Then
                        strInput = Mid$(strInput, lngPosition + 2)
                    Else
                        Exit Do
                    End If
                Loop
                If Len(strInput) = 0 Then
                    MsgBox "No source code was found in the input file " & GetFileName(strInputFile), vbCritical
                    Exit Function
                End If
            End If
        End If

        'reset a flag used by the tokenizer, to false
        BlnNextLineIsComment = False
        
        'convert code
        If ConvertCode(strInput) Then
            'output the converted code to file
            With Html
                lngListCount = .ListCount
                If lngListCount > 0 Then
                    If FileExist(strOutputFile) Then
                        FileKill strOutputFile
                    End If
                    intFileNum = FreeFile
                    On Error Resume Next
                    Open strOutputFile For Output As #intFileNum
                    If Err = 0 Then
                        If blnAddPreview Then    'add header
                            Print #intFileNum, "<html><head><title>Source Code Preview</title></head><body>" & vbCrLf
                        End If
                        'print out each line
                        For lngLoopIndex = 0 To lngListCount - 1
                            Print #intFileNum, .List(lngLoopIndex)
                        Next lngLoopIndex
                        If blnAddPreview Then    'add footer
                            Print #intFileNum, "</body></html>" & vbCrLf
                        End If
                        Close intFileNum
                    End If
                    On Error GoTo 0
                End If
            End With
            ConvertFile = True
        End If

    Else
        MsgBox "Invalid input file(" & GetFileName(strInputFile) & "). Could not read data or the file is empty. Please check the file path and try again.", vbCritical
    End If

End Function

'loads and parses a project file, calls convertfile() on each file which in turn calls convertcode()
Function ConvertProject(ByVal strProjectPath As String, ByVal strSaveFolder As String, ByVal blnOnlySource As Boolean, ByVal blnAddPreview As Boolean, Optional blnOverwrite As Boolean) As Boolean

    Dim NewProject      As VB_Project    'used to clear any current project

    Dim intFileNum      As Integer   'file number
    Dim strFileLine     As String    'current line from file
    Dim strSaveFile     As String    'file to save to
    Dim strInputPath    As String    'file to read from
    Dim lngCurTime      As Long      'current time
    Dim strBaseDir      As String    'base project folder

    'check valid project path
    If (Not FileExist(strProjectPath)) Or (Len(strProjectPath) = 0) Or (Not FileCanRead(strProjectPath)) Then
        MsgBox "Invalid Input Project File (" & GetFileName(strProjectPath) & "). Please check the file path and try again.", vbCritical
    Else

        Project = NewProject
        With Project
            .SaveDir = strSaveFolder
            If Not FolderExist(strSaveFolder) Then    'no output folder
                If MsgBox("The output folder (" & strSaveFolder & ") does not exist. Would you like it to be auto created?", vbCritical + vbYesNo) = vbYes Then
                    'makedeepdir() makes sure all folders from root (eg C:\) down exist
                    MakeDeepDir strSaveFolder
                Else
                    MsgBox "No Output Folder. Conversion Aborted.", vbCritical
                    Exit Function
                End If
            End If

            'grab our working proj dir from project path
            strBaseDir = GetFolderpath(strProjectPath)
            .BaseDir = strBaseDir
        End With
        
        TotalLines = 0
        TotalFiles = 0
        TotalTime = 0
        ProjectName = ""
        
        With frmMain
            If Not blnCmdMode Then
                .lstFiles.Clear
            End If


            intFileNum = FreeFile
            On Error Resume Next    'trap opening files for errors
            Open strProjectPath For Input As #intFileNum
            If Err = 0 Then    'all good
                On Error GoTo 0    'turn of error handling
                Do While Not EOF(intFileNum)
                    Line Input #intFileNum, strFileLine
                    If Len(strFileLine) > 0 Then

                        Select Case True
                            Case (Left$(strFileLine, 5) = "Name=")
                                ProjectName = Replace$(Mid$(strFileLine, InStr(strFileLine, "=") + 1), """", "")
                                AddProjectProperty strFileLine 'add to 'Project' UDT

                            Case (Left$(strFileLine, 5) = "Form="), _
                                    (Left$(strFileLine, 12) = "UserControl="), _
                                    (Left$(strFileLine, 13) = "PropertyPage="), _
                                    (Left$(strFileLine, 9) = "Designer="), _
                                    (Left$(strFileLine, 7) = "Module="), _
                                    (Left$(strFileLine, 6) = "Class=")
                                'todo: add webclass flag

                                'grab input file from project file
                                strInputPath = AbsoluteFromRelative(strBaseDir, GetVBPath(strFileLine))    'see GetVBPath() function, see modFile.bas for AbsoluteFromRelative()
                                'create a new save path from input file by adding a different extension
                                strSaveFile = strSaveFolder & "\" & CreateHtmlPath(GetFileName(strInputPath))    'see modFile for function GetFileName()
                                
                                If FileExist(strSaveFile) Then
                                    If Not blnOverwrite Then
                                        If MsgBox("The file " & GetFileName(strSaveFile) & " already exists! Do you wish to continue and overwrite any data already in the file?", vbCritical + vbYesNo) = vbNo Then
                                            Close #intFileNum
                                            Exit Function
                                        End If
                                     End If
                                End If
                                
                                
                                'save current time
                                lngCurTime = timeGetTime()    'api call
                                If Not blnCmdMode Then    'dont access frmMain in commandline mode
                                    .lblConverting = "Converting " & GetFileName(strInputPath) & "..."
                                End If

                                'convert code from file
                                If ConvertFile(strInputPath, strSaveFile, blnOnlySource, blnAddPreview, blnOverwrite) Then
                                    TotalTime = TotalTime + (timeGetTime - lngCurTime)
                                    TotalFiles = TotalFiles + 1
                                    'add to list box
                                    If Not blnCmdMode Then
                                        .lstFiles.AddItem strSaveFile
                                        .lstFiles.ListIndex = .lstFiles.NewIndex
                                        DoEvents
                                    End If
                                Else
                                    If MsgBox("Conversion of file (" & GetFileName(strInputPath) & ") in the project " & ProjectName & " failed. Do you want to continue with the rest of the files?", vbCritical + vbYesNo) = vbNo Then
                                        blnCancelled = True
                                        Close #intFileNum
                                        Exit Function
                                    End If
                                End If
                                
                                DoEvents
                               
                            Case Else
                                  AddProjectProperty strFileLine 'add to 'Project' UDT
                       
                       End Select
                    End If
                    If blnCancelled Then    'did we hit cancel?
                        Close #intFileNum
                        Exit Function
                    End If
                Loop
                Close #intFileNum
                With Project
                    .Initialized = True
                    .Other.Add CStr(TotalLines), "LineCount"
                    .Other.Add CStr(TotalFiles), "FileCount"
                End With
                ConvertProject = True
            Else
                Close 'close all/any open files
                MsgBox "Could not read input project file (" & strProjectPath & ")" & vbCrLf & vbCrLf & "[ debug: " & Err.Description & " ]", vbCritical
            End If   '/Err = 0/
        End With     '/frmMain/
    End If           '/Not FileExist(strProjectPath)/


End Function
