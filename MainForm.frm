VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VTH [vb to html]"
   ClientHeight    =   5970
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9810
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   9810
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "Vote For VTH"
      Height          =   375
      Left            =   120
      TabIndex        =   57
      Top             =   3180
      Width           =   1455
   End
   Begin VB.CommandButton cmdReadMe 
      Appearance      =   0  'Flat
      Caption         =   "View ReadMe"
      Height          =   375
      Left            =   120
      TabIndex        =   56
      Top             =   2340
      Width           =   1455
   End
   Begin VB.CommandButton cmdAbout 
      Appearance      =   0  'Flat
      Caption         =   "About VTH..."
      Height          =   375
      Left            =   120
      TabIndex        =   55
      Top             =   2760
      Width           =   1455
   End
   Begin VB.PictureBox picOverallProgress 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   3180
      ScaleHeight     =   195
      ScaleWidth      =   6495
      TabIndex        =   13
      Top             =   5520
      Width           =   6555
      Begin VB.Label lblOverallStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   6495
      End
      Begin VB.Label lblOverallProgress 
         BackColor       =   &H000000FF&
         Height          =   315
         Left            =   -45
         TabIndex        =   14
         Top             =   0
         Width           =   1035
      End
   End
   Begin VB.CommandButton cmdAction 
      Appearance      =   0  'Flat
      Caption         =   "Convert Project"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Top             =   1500
      Width           =   1455
   End
   Begin VB.CommandButton cmdAction 
      Appearance      =   0  'Flat
      Caption         =   "Convert File"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton cmdOptions 
      Appearance      =   0  'Flat
      Caption         =   "Color Options"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdAction 
      Appearance      =   0  'Flat
      Caption         =   "Convert Text"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   660
      Width           =   1455
   End
   Begin VB.Frame fraMain 
      Caption         =   "Convert Text"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   60
      Width           =   8055
      Begin VB.CommandButton cmdFileAction 
         Caption         =   "View Source"
         Height          =   375
         Index           =   2
         Left            =   2340
         TabIndex        =   54
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CommandButton cmdFileAction 
         Caption         =   "Save Source"
         Height          =   375
         Index           =   3
         Left            =   3600
         TabIndex        =   49
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CommandButton cmdFileAction 
         Caption         =   "Copy Source"
         Height          =   375
         Index           =   1
         Left            =   1080
         TabIndex        =   48
         Top             =   4920
         Width           =   1215
      End
      Begin VB.CommandButton cmdFileAction 
         Caption         =   "Preview In Browser"
         Height          =   375
         Index           =   4
         Left            =   6300
         TabIndex        =   17
         Top             =   4920
         Width           =   1635
      End
      Begin VB.CommandButton cmdFileAction 
         Caption         =   "Clear"
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   4920
         Width           =   975
      End
      Begin VB.CommandButton cmdConvert 
         Caption         =   "Convert"
         Height          =   375
         Left            =   6360
         TabIndex        =   10
         Top             =   1740
         Width           =   1575
      End
      Begin VB.TextBox txtInput 
         Height          =   1215
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   480
         Width           =   7935
      End
      Begin SHDocVwCtl.WebBrowser ctlWebBrowser 
         Height          =   2715
         Left            =   60
         TabIndex        =   12
         Top             =   2160
         Width           =   7875
         ExtentX         =   13891
         ExtentY         =   4789
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
      Begin VB.Label Label1 
         Caption         =   "Output:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1920
         Width           =   1755
      End
      Begin VB.Label Label1 
         Caption         =   "Input:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1755
      End
   End
   Begin VB.Frame fraMain 
      Caption         =   "Convert Project"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      Index           =   2
      Left            =   1680
      TabIndex        =   2
      Top             =   60
      Width           =   8055
      Begin VB.CommandButton cmdCreateIndex 
         Caption         =   "Create Master Index"
         Height          =   375
         Left            =   5640
         TabIndex        =   52
         Top             =   2160
         Width           =   1875
      End
      Begin VB.ListBox lstFiles 
         Height          =   1620
         Left            =   120
         TabIndex        =   46
         Top             =   3240
         Width           =   7815
      End
      Begin VB.TextBox txtFilePath 
         Height          =   375
         Index           =   2
         Left            =   1320
         TabIndex        =   43
         Top             =   240
         Width           =   6195
      End
      Begin VB.CommandButton cmdFileBrowse 
         Caption         =   "..."
         Height          =   375
         Index           =   2
         Left            =   7560
         TabIndex        =   42
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton cmdConvertProject 
         Caption         =   "Convert"
         Height          =   375
         Left            =   5640
         TabIndex        =   41
         Top             =   1320
         Width           =   1875
      End
      Begin VB.TextBox txtFilePath 
         Height          =   375
         Index           =   3
         Left            =   1320
         TabIndex        =   40
         Top             =   720
         Width           =   6195
      End
      Begin VB.CommandButton cmdFileBrowse 
         Caption         =   "..."
         Height          =   375
         Index           =   3
         Left            =   7560
         TabIndex        =   39
         Top             =   720
         Width           =   375
      End
      Begin VB.Frame fraOptions 
         Caption         =   "Options"
         Height          =   1635
         Index           =   1
         Left            =   180
         TabIndex        =   35
         Top             =   1200
         Width           =   5295
         Begin VB.CheckBox chkOverWrite 
            Caption         =   "Over Write Existing Files Without Prompting"
            Height          =   195
            Left            =   180
            TabIndex        =   50
            Top             =   1260
            Width           =   3735
         End
         Begin VB.OptionButton optConvertAll 
            Caption         =   "Convert Entire File Including Attributes And Properties"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   38
            Top             =   240
            Width           =   4155
         End
         Begin VB.OptionButton optOnlySource 
            Caption         =   "Convert Only Source Code"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   37
            Top             =   540
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.CheckBox chkAddPreview 
            Caption         =   "Output html source as viewable page (adds header,body tags)"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   36
            Top             =   900
            Value           =   1  'Checked
            Width           =   4995
         End
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview Html Source"
         Height          =   375
         Index           =   2
         Left            =   4140
         TabIndex        =   34
         Top             =   4920
         Width           =   1875
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview Html Page"
         Height          =   375
         Index           =   3
         Left            =   6060
         TabIndex        =   33
         Top             =   4920
         Width           =   1875
      End
      Begin VB.CommandButton cmdOpenOutPutFolder 
         Caption         =   "Open Output Folder"
         Height          =   375
         Index           =   1
         Left            =   5640
         TabIndex        =   32
         Top             =   1740
         Width           =   1875
      End
      Begin VB.Label lblConverting 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   4980
         Width           =   7695
      End
      Begin VB.Label Label3 
         Caption         =   "Results File List:"
         Height          =   315
         Left            =   120
         TabIndex        =   47
         Top             =   3000
         Width           =   4395
      End
      Begin VB.Label lblFilePath 
         Caption         =   "Project FilePath:"
         Height          =   255
         Index           =   3
         Left            =   60
         TabIndex        =   45
         Top             =   300
         Width           =   1815
      End
      Begin VB.Label lblFilePath 
         Caption         =   "Output Folder:"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   44
         Top             =   780
         Width           =   1755
      End
   End
   Begin VB.Frame fraMain 
      Caption         =   "Convert File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5355
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   60
      Width           =   8055
      Begin VB.CommandButton cmdOpenOutPutFolder 
         Caption         =   "Open Output Folder"
         Height          =   375
         Index           =   0
         Left            =   5880
         TabIndex        =   30
         Top             =   3360
         Width           =   1875
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview Html Page"
         Height          =   375
         Index           =   1
         Left            =   5880
         TabIndex        =   29
         Top             =   2940
         Width           =   1875
      End
      Begin VB.CommandButton cmdPreview 
         Caption         =   "Preview Html Source"
         Height          =   375
         Index           =   0
         Left            =   5880
         TabIndex        =   28
         Top             =   2520
         Width           =   1875
      End
      Begin VB.Frame fraOptions 
         Caption         =   "Options"
         Height          =   1755
         Index           =   0
         Left            =   180
         TabIndex        =   25
         Top             =   1980
         Width           =   5295
         Begin VB.CheckBox chkAddPreview 
            Caption         =   "Output html source as viewable page (adds header,body tags)"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   31
            Top             =   1260
            Width           =   5055
         End
         Begin VB.OptionButton optOnlySource 
            Caption         =   "Convert Only Source Code"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   27
            Top             =   720
            Value           =   -1  'True
            Width           =   2535
         End
         Begin VB.OptionButton optConvertAll 
            Caption         =   "Convert Entire File Including Attributes And Properties"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   26
            Top             =   360
            Width           =   4155
         End
      End
      Begin VB.CommandButton cmdFileBrowse 
         Caption         =   "..."
         Height          =   375
         Index           =   1
         Left            =   7380
         TabIndex        =   24
         Top             =   1440
         Width           =   375
      End
      Begin VB.TextBox txtFilePath 
         Height          =   375
         Index           =   1
         Left            =   180
         TabIndex        =   23
         Top             =   1440
         Width           =   7095
      End
      Begin VB.CommandButton cmdConvertFile 
         Caption         =   "Convert"
         Height          =   375
         Left            =   5880
         TabIndex        =   21
         Top             =   2100
         Width           =   1875
      End
      Begin VB.CommandButton cmdFileBrowse 
         Caption         =   "..."
         Height          =   375
         Index           =   0
         Left            =   7380
         TabIndex        =   20
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox txtFilePath 
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   600
         Width           =   7095
      End
      Begin VB.Label lblFilePath 
         Caption         =   "Output File Path:"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   22
         Top             =   1140
         Width           =   1755
      End
      Begin VB.Label lblFilePath 
         Caption         =   "Input File Path:"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Top             =   300
         Width           =   1155
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "VTH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   60
      TabIndex        =   51
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image imgIcon 
      Height          =   435
      Left            =   120
      Top             =   120
      Width           =   435
   End
   Begin VB.Label lblProgress 
      BackStyle       =   0  'Transparent
      Caption         =   "Overall Progress"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   1860
      TabIndex        =   15
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   60
      Shape           =   2  'Oval
      Top             =   60
      Width           =   1575
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strLastConvertedFilePath  As String
Dim strLastSelectedFilePath   As String

Private Sub Command1_Click()
   MsgBox "Thanks for your vote!", vbExclamation
   ExecuteFile "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57295&lngWId=1"
End Sub

Private Sub Form_Load()

    Dim lngIndex      As Long
    Dim strAgreePath  As String

    On Error Resume Next    'bypass load errors

    'get saved settings
    For lngIndex = 0 To txtFilePath.Count - 1
        txtFilePath(lngIndex).Text = ReadString("files", "filepath" & CStr(lngIndex), "")
    Next
    chkAddPreview(0).Value = ReadNumber("options", "addpreview0", 1)
    chkAddPreview(1).Value = ReadNumber("options", "addpreview1", 1)
    If ReadNumber("options", "convertall0", 0) Then
        optConvertAll(0).Value = True
    End If
    If ReadNumber("options", "convertall1", 0) Then
        optConvertAll(1).Value = True
    End If
    chkOverWrite.Value = ReadNumber("options", "overwrite", 0)

    'set defaults
    Width = 9900
    Set imgIcon = Icon
    SetOverallProgressMax 0
    ctlWebBrowser.Navigate StartPage(imgIcon)
    fraMain(0).ZOrder

    If ReadNumber("settings", "firstrun", 0) = 0 Then
        MsgBox "Thank you for using VTH. Please read the included agreement before using this software or source code.", vbInformation
        WriteNumber "settings", "firstrun", 1
        Show
        DoEvents
        'show color options if this is first time
        frmOptions.Show , Me

        strAgreePath = strAppPath & "\agreement.txt"
        If Not FileExist(strAgreePath) Then
            FileSave StrConv(LoadResData("agree", "text"), vbUnicode), strAgreePath
        End If
        ExecuteFile strAgreePath
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim lngIndex As Long
    Dim strTempFile As String
    Dim frmGeneric As Form

    On Error Resume Next    'lets not fail while closing :(

    'save settings
    For lngIndex = 0 To txtFilePath.Count - 1
        WriteString "files", "filepath" & CStr(lngIndex), txtFilePath(lngIndex).Text
    Next
    WriteNumber "options", "addpreview0", chkAddPreview(0).Value
    WriteNumber "options", "addpreview1", chkAddPreview(1).Value
    'Abs() converts -1 to 1
    WriteNumber "options", "convertall0", Abs(CInt(optConvertAll(0).Value))
    WriteNumber "options", "convertall1", Abs(CInt(optConvertAll(1).Value))
    WriteNumber "options", "overwrite", chkOverWrite.Value

    'clean up temp files
    strTempFile = Dir$(strAppPath & "\*.txt")
    Do While Len(strTempFile) > 0
        If InStr(strTempFile, "fileprev") > 0 Or (InStr(strTempFile, "fileprev") > 0 And InStr(strTempFile, "[") > 0) Then
            FileKill strAppPath & "\" & strTempFile
        End If
        strTempFile = Dir$
    Loop

    strTempFile = Dir$(strAppPath & "\*.html")
    Do While Len(strTempFile) > 0
        If InStr(strTempFile, "fileprev") > 0 Or (InStr(strTempFile, "fileprev") > 0 And InStr(strTempFile, "[") > 0) Then
            FileKill strAppPath & "\" & strTempFile
        End If
        strTempFile = Dir$
    Loop

    'all these files are dynamically re-created from
    'the res file or from code if they dont exist
    FileKill strAppPath & "\preview.html"
    FileKill strAppPath & "\progress.html"
    FileKill strAppPath & "\welcome.html"
    FileKill strAppPath & "\error.html"
    FileKill strAppPath & "\blank.html"
    FileKill strAppPath & "\progbar.gif"
    FileKill strAppPath & "\icon.ico"

    'close any open forms that isnt this one
    For Each frmGeneric In Forms
        If frmGeneric.hWnd <> Me.hWnd Then
            Unload frmGeneric
            Set frmGeneric = Nothing
        End If
    Next frmGeneric

End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show , Me
End Sub

'change frames
Private Sub cmdAction_Click(Index As Integer)
    fraMain(Index).ZOrder
End Sub

Private Sub cmdCreateIndex_Click()
    If Project.Initialized Then
        frmCreateIndex.Show , Me
    Else
        MsgBox "You must first convert a project before an index file can be created.", vbCritical
    End If
End Sub

'convert source text to html
Private Sub cmdConvert_Click()

    If cmdConvert.Caption = "Cancel" Then
        blnCancelled = True
        cmdConvert.Caption = "Convert"
    Else
        'make sure there is code
        If Len(txtInput.Text) > 0 Then

            ctlWebBrowser.Navigate ProgressPagePath()    'see function ProgressPagePath() in modCode.bas
            DoEvents

            'reset a flag used by the tokenizer to false
            BlnNextLineIsComment = False
            cmdConvert.Caption = "Cancel"
            strLastConvertedFilePath = ""
            blnCancelled = False
            TotalLines = 0

            If ConvertCode(txtInput.Text) Then
                If Not blnCancelled Then
                    strLastConvertedFilePath = MakeHtmlPage() 'see function MakeHtmlPage() in modCode.bas
                    ctlWebBrowser.Navigate strLastConvertedFilePath
                Else
                    ctlWebBrowser.Navigate BlankPage()
                End If
            Else
                ctlWebBrowser.Navigate ErrorPage()    'see function ErrorPage() in modCode.bas
            End If

        Else
            MsgBox "Please paste in some code to convert!", vbCritical
        End If
        cmdConvert.Caption = "Convert"
        blnCancelled = False
    End If

End Sub

'convert a single file
Private Sub cmdConvertFile_Click()

    If cmdConvertFile.Caption = "Cancel" Then
        If MsgBox("Are you sure you want to cancel the file conversion?", vbQuestion + vbYesNo) = vbYes Then
            cmdConvertFile.Caption = "Convert"
            blnCancelled = True
        End If
    Else
        cmdConvertFile.Caption = "Cancel"
        blnCancelled = False
        TotalLines = 0
        If ConvertFile(txtFilePath(0), txtFilePath(1), optOnlySource(0).Value, chkAddPreview(0).Value) Then
            If Not blnCancelled Then
                strLastConvertedFilePath = txtFilePath(1)
                MsgBox "Converted " & CStr(TotalLines) & " lines in " & CalculateTime(timeGetTime() - StartTime) & " to file " & txtFilePath(1), vbInformation
            Else
               MsgBox "Conversion was cancelled.", vbCritical
            End If
        Else
            MsgBox "Conversion was unsuccessful! Please check all input and try again.", vbCritical
        End If
        cmdConvertFile.Caption = "Convert"
        blnCancelled = False
    End If

End Sub

Private Sub cmdConvertProject_Click()

    If cmdConvertProject.Caption = "Cancel" Then
        If MsgBox("Are you sure you want to cancel the project conversion?", vbQuestion + vbYesNo) = vbYes Then
            cmdConvertProject.Caption = "Convert"
            blnCancelled = True
        End If
    Else
        cmdConvertProject.Caption = "Cancel"
        blnCancelled = False
        
        'convert the project
        If ConvertProject(txtFilePath(2), txtFilePath(3), optOnlySource(1).Value, chkAddPreview(1).Value, chkOverWrite.Value) Then
            MsgBox "Conversion of the project " & ProjectName & " is complete. A total of " & CStr(TotalLines) & " lines in " & CStr(TotalFiles) & " files, were converted to html source code in " & CalculateTime(TotalTime) & " to the folder " & txtFilePath(3) & vbCrLf & vbCrLf & "Double-Click on a file in the list to preview it in your browser, Or select it in the list and choose one of the two viewing option buttons.", vbInformation
        Else
            If blnCancelled Then    'did we hit cancel?
                MsgBox "Conversion of the project " & ProjectName & " was cancelled. A total of " & CStr(TotalLines) & " lines in " & CStr(TotalFiles) & " files, were converted to html source code in " & CalculateTime(TotalTime) & " to the folder " & txtFilePath(3) & vbCrLf & vbCrLf & "Double-Click on a file in the list to preview it in your browser, Or select it in the list and choose one of the two viewing option buttons.", vbInformation
            Else
                MsgBox "Conversion Failed.", vbCritical
            End If
        End If

        lblOverallStatus = "Converted " & CStr(TotalFiles) & " Files With " & CStr(TotalLines) & " Lines In " & CalculateTime(TotalTime)
        cmdConvertProject.Caption = "Convert"
        blnCancelled = False
        lblConverting = ""
    End If
    
End Sub

Private Sub cmdFileAction_Click(Index As Integer)

    Dim strSavePath As String

    Select Case Index
        Case 0       'clear
            strLastConvertedFilePath = ""
            ctlWebBrowser.Navigate BlankPage
            Exit Sub

        Case 1       'copy source
            If FileExist(strLastConvertedFilePath) Then
                Clipboard.Clear
                On Error Resume Next    'is there a thing as to much data?
                Clipboard.SetText FileToString(strLastConvertedFilePath)
                If Err = 0 Then
                    MsgBox "Source code has been copied to clipboard.", vbInformation
                Else
                    MsgBox Err.Description, vbCritical
                End If
                Exit Sub
            End If

        Case 2       'view source
            If FileExist(strLastConvertedFilePath) Then
                strSavePath = SafeFilePath(strAppPath & "\fileprev.txt")
                FileSave FileToString(strLastConvertedFilePath), strSavePath
                ExecuteFile strSavePath
                Exit Sub
            End If

        Case 3       'save to file
            If Not Html Is Nothing Then
                strSavePath = ShowSave("Text Files|*.txt|Html Files|*.html", "", "Save to File", "html")
                If strSavePath <> "" Then
                    MakeHtmlPage strSavePath    'will write source to text or html file
                    MsgBox "File Saved To " & strSavePath, vbInformation
                End If
                Exit Sub
            End If

        Case 4       'preview in browser
            strSavePath = strAppPath & "\preview.html"
            If FileExist(strSavePath) Then
                ExecuteFile strSavePath
                Exit Sub
            End If

    End Select

    MsgBox "No input data is available.", vbCritical

End Sub

'main sub for choosing a file or folder path
Private Sub cmdFileBrowse_Click(Index As Integer)

    Dim FilePath As String

    Select Case Index
        Case 0
            FilePath = ShowOpen("Visual Basic Files|*.FRM;*.BAS;*.CTL;*.CLS;*.DSR;*.DOB;*.PAG", "", "Open file...", "bas", CurDir$)
        Case 1
            FilePath = ShowSave("Text Files|*.txt|Html Files|*.html", "", "Save file...", "txt", CurDir$)
        Case 2
            FilePath = ShowOpen("Visual Basic Projects|*.VBP", "", "Open file...", "prj", CurDir$)
        Case 3
            FilePath = ShowBrowse("Select output folder...", CurDir$)
    End Select

    If Len(FilePath) > 0 Then
        txtFilePath(Index) = FilePath
    End If

End Sub

'one sub to bind them, one sub to rule them all ...
Private Sub cmdPreview_Click(Index As Integer)

    Dim NewFilePath As String
    Dim SourceFilePath As String

    If Index < 2 Then 'single file frame
        SourceFilePath = strLastConvertedFilePath
    Else             'project file frame
        SourceFilePath = strLastSelectedFilePath
    End If

    If FileExist(SourceFilePath) Then

        If (Index Mod 2) = 0 Then    'source
            NewFilePath = strAppPath & "\fileprev.txt"
        Else         'html
            NewFilePath = strAppPath & "\fileprev.html"
        End If

        If SourceFilePath = NewFilePath Then
            'safefilepath() will guarauntee that the new file path wont exist
            'see modFile.bas for SafeFilePath()
            NewFilePath = SafeFilePath(NewFilePath)
        Else
            If FileExist(NewFilePath) Then
                FileKill NewFilePath
            End If
        End If

        If Not FileHasHeader(SourceFilePath) Then
            'add with html header
            FileSave "<html><head><title>Source Code Preview</title></head><body>" & vbCrLf & FileToString(SourceFilePath) & "</body></html>" & vbCrLf, NewFilePath
        Else
            'already has header, just copy over
            On Error Resume Next    'filecopy is notorious for errors
            FileCopy SourceFilePath, NewFilePath
            If Err Then
                MsgBox Err.Description, vbCritical
                Exit Sub
            End If
            On Error GoTo 0
        End If

        ExecuteFile NewFilePath

    Else
        MsgBox "No file found to open.", vbCritical
    End If
End Sub

Private Sub cmdReadMe_Click()

    Dim strReadPath As String

    strReadPath = strAppPath & "\readme.txt"
    If Not FileExist(strReadPath) Then
        FileSave StrConv(LoadResData("read", "text"), vbUnicode), strReadPath
    End If
    ExecuteFile strReadPath

End Sub

Private Sub cmdOpenOutputFolder_Click(Index As Integer)

    Dim strFolderPath As String, strFilePath As String

    If Index = 0 Then 'file frame
        strFilePath = strLastConvertedFilePath
    Else             'project frame
        strFilePath = strLastSelectedFilePath
    End If

    If Not FileExist(strFilePath) Then
        'try the one in the text box
        strFolderPath = txtFilePath((Index * 2) + 1)
    Else
        'get the real folder path to last processed file path
        strFolderPath = GetFolderpath(strFilePath)
    End If

    If FolderExist(strFolderPath) Then
        'executefile() also opens folders ...
        ExecuteFile strFolderPath
    Else
        MsgBox "Unable to open folder.", vbCritical
    End If

End Sub

Private Sub cmdOptions_Click()
    frmOptions.Show , Me
End Sub

Private Sub lstFiles_Click()
    strLastSelectedFilePath = lstFiles.Text
End Sub

Private Sub lstFiles_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    strLastSelectedFilePath = lstFiles.Text
End Sub

Private Sub lstFiles_DblClick()
    Call cmdPreview_Click(3)
End Sub

'**************************************************
'Helper Functions
'**************************************************

'these progress functions are wrappers for the label
'and picture used as a progress bar
Sub SetOverallProgress(ByVal Progress As Long)
    On Error Resume Next
    If Progress > picOverallProgress.ScaleWidth Then
        SetOverallProgressMax Progress
    End If
    lblOverallProgress.Width = Progress
    On Error GoTo 0
End Sub

Sub SetOverallProgressMax(ByVal ProgressMax As Long)
    On Error Resume Next
    lblOverallProgress.Width = 0
    picOverallProgress.ScaleWidth = ProgressMax
    On Error GoTo 0
End Sub
