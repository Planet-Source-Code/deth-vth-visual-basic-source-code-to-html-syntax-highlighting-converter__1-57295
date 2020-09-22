VERSION 5.00
Begin VB.Form frmCreateIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create Master Index Html Page"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5385
   Icon            =   "frmCreateIndex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5385
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdUseProject 
      Caption         =   "Use <Project Name>.html"
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   2160
      Width           =   2235
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4080
      TabIndex        =   8
      Top             =   2160
      Width           =   1155
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   2820
      TabIndex        =   7
      Top             =   2160
      Width           =   1155
   End
   Begin VB.CheckBox chkOpenIndex 
      Caption         =   "Open Index Page When Completed"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   1620
      Width           =   3015
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Index           =   2
      Left            =   1560
      TabIndex        =   5
      Top             =   1080
      Width           =   3675
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Index           =   1
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   3675
   End
   Begin VB.TextBox txtInput 
      Height          =   375
      Index           =   0
      Left            =   1560
      TabIndex        =   2
      Top             =   120
      Width           =   3675
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Index Page Name:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1140
      Width           =   1575
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Author:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Name:"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   1155
   End
End
Attribute VB_Name = "frmCreateIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim blnOpenedPage As Boolean

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCreate_Click()

    'remove any leading/trailing spaces
    If Len(txtInput(2)) > 0 Then
        txtInput(2) = Trim$(txtInput(2))
    End If

    If Len(txtInput(2)) < 4 Then
        MsgBox "Please enter a valid html index file name, with the extension .html", vbCritical
        Exit Sub
    End If

    'make sure we have a valid file extension
    If Not (InStr(txtInput(2), ".html") = Len(txtInput(2)) - 4) Then
        txtInput(2) = txtInput(2) & ".html"
    End If

    On Error Resume Next
    With Project

        'check overwrite
        If FileExist(Project.SaveDir & "\" & txtInput(2)) Then
            If frmMain.chkOverWrite.Value = 0 Then    'prompt for overwrite
                If MsgBox("This master index page '" & txtInput(2) & "' already exists, do you wish to overwrite?", vbCritical + vbYesNo) = vbNo Then
                    Exit Sub
                End If
            End If
        End If

        'set name
        If PropertyExist("Name") Then
            .Other("Name") = txtInput(0)
        Else
            .Other.Add txtInput(0), "Name"
        End If
        If .Other("Name") = "" Then .Other("Name") = "Project"

        'set author
        If PropertyExist("VersionCompanyName") Then
            .Other("VersionCompanyName") = txtInput(1)
        Else
            .Other.Add txtInput(1), "VersionCompanyName"
        End If

    End With
    On Error GoTo 0

    Hide
    DoEvents

    'pop back to the top (for an irratating flaw in vb when a form is shown in child mode)
    If (Not (frmMain Is Nothing)) Then
        If frmMain.WindowState <> vbMinimized Then
            frmMain.Show
        End If
    End If

    'create index page
    CreateMasterIndexPage txtInput(2), chkOpenIndex.Value
    'save option
    WriteNumber "options", "openindex", chkOpenIndex.Value
    'set flag for unload sub
    blnOpenedPage = (chkOpenIndex.Value = 1)

    Unload Me

End Sub

Private Sub cmdUseProject_Click()

    Dim strInputName As String

    If PropertyExist("Name") Then
        txtInput(2) = LCase$(Project.Other("Name")) & ".html"
    Else
        If PropertyExist("Title") Then
          txtInput(2) = LCase$(Project.Other("Title")) & ".html"
        Else
            'prompt for a name
            strInputName = InputBox$("Project title not found. Please enter a project name to use.", "Error", "Project1")
            If StrPtr(strInputName) = 0 Then
                'ok were done trying
                txtInput(2) = "index.html"
            Else
                If Len(strInputName) = 0 Then    '0 length!
                    txtInput(2) = "index.html"    'use default
                Else
                    strInputName = LCase$(Trim$(strInputName))
                    'check for extension
                    If Not (InStr(strInputName, ".html") = Len(strInputName) - 4) Then
                        strInputName = strInputName & ".html"
                    End If
                    txtInput(2) = strInputName
                End If
            End If
        End If
    End If

End Sub

Private Sub Form_Load()
    Icon = frmMain.Icon

    On Error Resume Next
    With Project
        If .Initialized Then
            txtInput(0) = .Other("Name")
            txtInput(1) = .Other("VersionCompanyName")
            txtInput(2) = "index.html"
        End If
    End With

    chkOpenIndex.Value = ReadNumber("options", "openindex", 0)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If Not blnOpenedPage Then    'we dont want to jump on top
        '(for an irratating flaw in vb when a form is shown in child mode)
        If (Not (frmMain Is Nothing)) Then
            If frmMain.WindowState <> vbMinimized Then
                frmMain.Show
            End If
        End If
    End If

End Sub

