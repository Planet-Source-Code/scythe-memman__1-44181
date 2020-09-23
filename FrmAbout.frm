VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "About MemMan"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2925
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   2925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   255
      Left            =   1140
      TabIndex        =   4
      Top             =   1620
      Width           =   675
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      Caption         =   "www.scythe-tools.de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      MouseIcon       =   "FrmAbout.frx":0000
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   3
      Top             =   1260
      Width           =   2895
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      Caption         =   "scythe@cablenet.de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   0
      MouseIcon       =   "FrmAbout.frx":0152
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   2
      Top             =   900
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      Caption         =   "© 2003 Scythe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   540
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      Caption         =   "MemMan V1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   2835
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
 Unload Me
End Sub
Private Function HyperJump(ByVal URL As String) As Long
 HyperJump = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function

Private Sub Form_Load()
 SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub Label3_Click()
 HyperJump "mailto:" & Label3.Caption
End Sub

Private Sub Label4_Click()
 HyperJump Label4.Caption
End Sub
