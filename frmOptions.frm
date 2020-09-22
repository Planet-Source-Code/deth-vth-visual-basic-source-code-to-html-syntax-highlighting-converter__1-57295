VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Html Code Colors"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5700
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkAddLine 
      Caption         =   "Add Lines Between Subs And Functions"
      Height          =   255
      Left            =   1860
      TabIndex        =   30
      Top             =   3660
      Width           =   3435
   End
   Begin VB.CommandButton cmdDefaults 
      Caption         =   "Defaults"
      Height          =   375
      Left            =   240
      TabIndex        =   29
      Top             =   4140
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4380
      TabIndex        =   28
      Top             =   4140
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   27
      Top             =   4140
      Width           =   1095
   End
   Begin VB.Label lblPreview 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   8
      Left            =   2280
      TabIndex        =   26
      Top             =   3060
      Width           =   3300
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   8
      Left            =   1740
      TabIndex        =   25
      Top             =   3060
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BackGround Color:"
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   24
      Top             =   3120
      Width           =   1515
   End
   Begin VB.Label lblPreview 
      BackColor       =   &H00FFFFFF&
      Caption         =   "#Const #If #Else #End IF"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   405
      Index           =   7
      Left            =   2280
      TabIndex        =   23
      Top             =   2700
      Width           =   3300
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   7
      Left            =   1740
      TabIndex        =   22
      Top             =   2700
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Directive Color:"
      Height          =   195
      Index           =   7
      Left            =   120
      TabIndex        =   21
      Top             =   2760
      Width           =   1515
   End
   Begin VB.Label lblPreview 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1234567890"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Index           =   6
      Left            =   2280
      TabIndex        =   20
      Top             =   2340
      Width           =   3300
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   6
      Left            =   1740
      TabIndex        =   19
      Top             =   2340
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Number Color:"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   1515
   End
   Begin VB.Label lblPreview 
      BackColor       =   &H00FFFFFF&
      Caption         =   "<= + * ^ & % $ # - / > < >="
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Index           =   5
      Left            =   2280
      TabIndex        =   17
      Top             =   1980
      Width           =   3300
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   5
      Left            =   1740
      TabIndex        =   16
      Top             =   1980
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Operator Color:"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   2040
      Width           =   1515
   End
   Begin VB.Label lblPreview 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Form_Load()"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Index           =   4
      Left            =   2280
      TabIndex        =   14
      Top             =   1620
      Width           =   3300
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   4
      Left            =   1740
      TabIndex        =   13
      Top             =   1620
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Normal Color:"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   1515
   End
   Begin VB.Label lblPreview 
      BackColor       =   &H00FFFFFF&
      Caption         =   """this is a string"""
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Index           =   3
      Left            =   2280
      TabIndex        =   11
      Top             =   1260
      Width           =   3300
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   3
      Left            =   1740
      TabIndex        =   10
      Top             =   1260
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "String Color:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   1515
   End
   Begin VB.Label lblPreview 
      BackColor       =   &H00FFFFFF&
      Caption         =   "intFileNum FilePath "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Index           =   2
      Left            =   2280
      TabIndex        =   8
      Top             =   900
      Width           =   3300
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   2
      Left            =   1740
      TabIndex        =   7
      Top             =   900
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Identifier Color:"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1515
   End
   Begin VB.Label lblPreview 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Private Public Sub Function As"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   405
      Index           =   1
      Left            =   2280
      TabIndex        =   5
      Top             =   540
      Width           =   3300
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00400000&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   1
      Left            =   1740
      TabIndex        =   4
      Top             =   540
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Keyword Color:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1515
   End
   Begin VB.Label lblPreview 
      BackColor       =   &H00FFFFFF&
      Caption         =   "'this is a comment"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   405
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   180
      Width           =   3300
   End
   Begin VB.Label lblColor 
      BackColor       =   &H00008000&
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Index           =   0
      Left            =   1740
      TabIndex        =   1
      Top             =   180
      Width           =   375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment Color:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1515
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    Dim lngLoopIndex As Long

    'set color array to chosen color labels
    For lngLoopIndex = 0 To lblColor.Count - 1
        ColorArr(lngLoopIndex) = lblColor(lngLoopIndex).BackColor
        HexArr(lngLoopIndex) = vbQuote & HtmlColor(ColorArr(lngLoopIndex)) & vbQuote
    Next lngLoopIndex

    blnAddLine = (chkAddLine.Value > 0)
    
    SaveColorSettings
   
    Unload Me

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'load default colors
Private Sub cmdDefaults_Click()

    Dim lngLoopIndex As Long

    For lngLoopIndex = 0 To lblColor.Count - 1
        lblColor(lngLoopIndex).BackColor = DefColArr(lngLoopIndex)
        lblPreview(lngLoopIndex).ForeColor = DefColArr(lngLoopIndex)
        lblPreview(lngLoopIndex).BackColor = vbWhite
    Next lngLoopIndex

End Sub

Private Sub Form_Load()

    Dim lngLoopIndex As Long

    'by re-using the icon from the main form and deleting this one from the
    'property window, you can significantly reduce the size of your
    'compiled applications...
    Icon = frmMain.Icon

    'load the saved colors into labels
    For lngLoopIndex = 0 To lblColor.Count - 1
        lblColor(lngLoopIndex).BackColor = ColorArr(lngLoopIndex)
        lblPreview(lngLoopIndex).BackColor = ColorArr(8)
        lblPreview(lngLoopIndex).ForeColor = ColorArr(lngLoopIndex)
    Next lngLoopIndex
    
    chkAddLine.Value = Abs(CInt(blnAddLine))
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'this is my personal fight to make vb do what i want...
        If (Not (frmMain Is Nothing)) Then
            If frmMain.WindowState <> vbMinimized Then
                frmMain.Show
            End If
        End If
    
End Sub

Private Sub lblColor_Click(Index As Integer)
    'see helper function
    SelectNewColor Index
End Sub

Private Sub lblPreview_Click(Index As Integer)
    'see helper function
    SelectNewColor Index
End Sub

'helper function to choose new color
Sub SelectNewColor(ByVal Index As Integer)

    Dim ChosenColor As Long, lngLoopIndex As Long

    ChosenColor = lblColor(Index).BackColor 'default to current color
    ChosenColor = ShowColor(ChosenColor)    'show color picker
    If ChosenColor <> -1 Then 'new color chosen
        lblColor(Index).BackColor = ChosenColor
        lblPreview(Index).ForeColor = ChosenColor
        If Index = lblColor.Count - 1 Then
            For lngLoopIndex = 0 To lblColor.Count - 1
                lblPreview(lngLoopIndex).BackColor = ChosenColor
            Next lngLoopIndex
        End If
    End If

End Sub
