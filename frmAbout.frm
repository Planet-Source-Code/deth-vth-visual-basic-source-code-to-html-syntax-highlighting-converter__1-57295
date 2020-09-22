VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   2700
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   4995
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1863.588
   ScaleMode       =   0  'User
   ScaleWidth      =   4690.563
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   405
      Left            =   3300
      TabIndex        =   0
      Top             =   2160
      Width           =   1260
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "  Author: Lewis Miller"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1020
      TabIndex        =   4
      Top             =   1080
      Width           =   3885
   End
   Begin VB.Image imgIcon 
      Height          =   555
      Left            =   240
      Top             =   300
      Width           =   555
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   197.201
      X2              =   4338.419
      Y1              =   1397.691
      Y2              =   1397.691
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "App Description"
      ForeColor       =   &H00000000&
      Height          =   450
      Left            =   1020
      TabIndex        =   1
      Top             =   1440
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   225.372
      X2              =   4338.419
      Y1              =   1408.044
      Y2              =   1408.044
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1050
      TabIndex        =   3
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "VTH  (Vb To Html)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   480
      Index           =   1
      Left            =   1012
      TabIndex        =   5
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "VTH  (Vb To Html)"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   480
      Index           =   0
      Left            =   1050
      TabIndex        =   2
      Top             =   240
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    Icon = frmMain.Icon
    imgIcon = frmMain.Icon
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
    lblDescription.Caption = App.FileDescription
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'this is my personal fight to make vb do what i want...
        If (Not (frmMain Is Nothing)) Then
            If frmMain.WindowState <> vbMinimized Then
                frmMain.Show
            End If
        End If
End Sub
