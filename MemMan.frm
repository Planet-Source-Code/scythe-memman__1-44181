VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form FrmMemMan 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "MemMan"
   ClientHeight    =   2250
   ClientLeft      =   150
   ClientTop       =   675
   ClientWidth     =   1815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   150
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   121
   StartUpPosition =   3  'Windows-Standard
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   60
      TabIndex        =   4
      Top             =   1140
      Width           =   1755
      Begin VB.TextBox TxtMinRam 
         Height          =   285
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "10 %"
         ToolTipText     =   "Transparency in %"
         Top             =   660
         Width           =   435
      End
      Begin VB.VScrollBar VScrMinRam 
         Height          =   285
         Left            =   480
         Max             =   1
         Min             =   90
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   660
         Value           =   10
         Width           =   135
      End
      Begin VB.CheckBox ChkAuto 
         Caption         =   "Auto free Memory  "
         Height          =   375
         Left            =   60
         TabIndex        =   5
         Top             =   180
         Value           =   1  'Aktiviert
         Width           =   1635
      End
      Begin VB.Label Label1 
         Caption         =   "Minimum"
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   720
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1035
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   1755
      Begin VB.OptionButton OptPerc 
         Caption         =   "Show as Percent"
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton OptGraph 
         Caption         =   "Show as Graph"
         Height          =   315
         Left            =   60
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1140
      Top             =   2520
   End
   Begin VB.PictureBox Pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      Height          =   240
      Left            =   840
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   0
      Top             =   2520
      Width           =   240
   End
   Begin ComctlLib.ImageList IL1 
      Left            =   240
      Top             =   2460
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MemMan.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu MnuPop 
      Caption         =   "Pop"
      Visible         =   0   'False
      Begin VB.Menu FreeMem 
         Caption         =   "Free Memory"
      End
      Begin VB.Menu MnuSettings 
         Caption         =   "Settings"
      End
      Begin VB.Menu MnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu MnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEnd 
         Caption         =   "End"
      End
   End
   Begin VB.Menu MnuMnu 
      Caption         =   "Menu"
      Begin VB.Menu MnuHide 
         Caption         =   "Hide"
      End
      Begin VB.Menu MnuLine3 
         Caption         =   "-"
      End
      Begin VB.Menu MnuEnd2 
         Caption         =   "End"
      End
   End
   Begin VB.Menu MnuAbout2 
      Caption         =   "About"
   End
End
Attribute VB_Name = "FrmMemMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'MemMan
'by Scythe
'scythe@cablenet
'www.scythe-tools.de

Private Declare Sub GlobalMemoryStatus Lib "KERNEL32" (lpBuffer As MemoryStatus)

Private Type MemoryStatus
 dwLength As Long
 dwMemoryLoad As Long
 dwTotalPhys As Long
 dwAvailPhys As Long
 dwTotalPageFile As Long
 dwAvailPageFile As Long
 dwTotalVirtual As Long
 dwAvailVirtual As Long
End Type

Dim ShowMode As Long
Dim ToTray As Boolean
Dim LastTry As Boolean

Private Sub Form_Load()
 'Load the Settings from registry
 ChkAuto.Value = Val(GetSetting(appname:="MemMan", section:="Settings", Key:="Auto", Default:="1"))
 VScrMinRam.Value = Val(GetSetting(appname:="MemMan", section:="Settings", Key:="Min", Default:="10"))
 TxtMinRam = VScrMinRam & " %"
 ShowMode = Val(GetSetting(appname:="MemMan", section:="Settings", Key:="View", Default:="0"))
 If ShowMode = 1 Then OptGraph.Value = True
 
 'Add the tray Icon
 AddToTray Me, MnuPop
 'Hide this form
 Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 'UloadMode = 0 if X was pressed
 If UnloadMode = 0 Then
  Me.Hide
  Cancel = True
  ToTray = False
 Else
  'Quti the program
  RemoveFromTray
 End If
End Sub

Private Sub Form_Resize()
 'Set the window on top
 SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE + SWP_NOSIZE
End Sub

Private Sub ChkAuto_Click()
 'Automatic FreeMem is selected or not
 'so en/disable the other needed controls
 Dim Tmp As Boolean
 If ChkAuto.Value = 1 Then Tmp = True
 TxtMinRam.Enabled = Tmp
 VScrMinRam.Enabled = Tmp
End Sub

'Change the Minimum Ramsize
Private Sub VScrMinRam_Change()
 TxtMinRam = VScrMinRam & " %"
End Sub

'Show the About Dialog
Private Sub MnuAbout_Click()
 FrmAbout.Show 1
End Sub
Private Sub MnuAbout2_Click()
 MnuAbout_Click
End Sub

'Quit the program
Private Sub MnuEnd_Click()
 'Save settings to Registry
 SaveSetting appname:="MemMan", section:="Settings", Key:="Auto", setting:=ChkAuto.Value
 SaveSetting appname:="MemMan", section:="Settings", Key:="Min", setting:=VScrMinRam.Value
 SaveSetting appname:="MemMan", section:="Settings", Key:="View", setting:=ShowMode
 Unload Me
End Sub
Private Sub MnuEnd2_Click()
 MnuEnd_Click
End Sub

'Hide (move to tray)
Private Sub MnuHide_Click()
Me.Hide
End Sub

'Show the settings form)
Private Sub MnuSettings_Click()
 Me.Show
End Sub

'Select graphical View
Private Sub OptGraph_Click()
 If OptGraph.Value = True Then ShowMode = 1
End Sub

'Select Percent View
Private Sub OptPerc_Click()
 If OptPerc.Value = True Then ShowMode = 0
End Sub

'The Display Routine
Private Sub Timer1_Timer()
 Dim MemDat As MemoryStatus
 Dim Tmp As String
 Dim x As Long
 Dim y As Long
 Dim Z As Long
 
 'Get Memory data
 MemDat.dwLength = Len(MemDat)
 GlobalMemoryStatus MemDat
 
 'Fill needed Variables with the Data
 Z = CLng(MemDat.dwAvailPhys / (MemDat.dwTotalPhys / 100))
 Tmp = Str(Z)

 'Clear Picture
 Pic.Picture = LoadPicture()

 'Sohwmode 0 is Percent view
 If ShowMode = 0 Then
  'Set Position and Print Text to Pic
  Pic.CurrentX = (Pic.Width - Pic.TextWidth(Tmp)) / 2
  Pic.CurrentY = (Pic.Height - Pic.TextHeight(Tmp)) / 2
  Pic.Print Tmp
 Else
  'Show Graphical Memory stats
  'get coordinats and color. Drw line to Pic
  x = Z * 2.5
  y = 15 - Z * 0.15
  Pic.Line (14, 16)-(1, y), RGB(255 - x, 0, x), BF
 End If
 'Store the whole thing in our ImageList
 IL1.ListImages.Add 2, , Pic.Image
 'Get it back as Icon
 Pic.Picture = IL1.ListImages(2).ExtractIcon
 'Delete new createt Image from list
 IL1.ListImages.Remove (2)
 
 'Now show the new Icon
 UpdateIcon
 'Set Tooltip
 SetTrayTip CLng(MemDat.dwAvailPhys / 1048576) & " MB of " & CLng(MemDat.dwTotalPhys / 1048576) & " MB free (" & Z & "%)"
 
 'Is Automatic FreeMem selected
 If ChkAuto.Value = 1 Then
  'Not enough ram ?
  If Z <= VScrMinRam.Value Then
   'LastTry is True if we moved ram with last timer call
   'and cant get what the user set to minimum RAM
   If LastTry = True Then
    'So set minimum ram down
    'If we dont do this the pc will try to free ram every time
    VScrMinRam.Value = Z - 1
    TxtMinRam = VScrMinRam & " %"
    LastTry = False
    Exit Sub
   Else
    'Else we free the ram now
    FreeMem_Click
    LastTry = True
   End If
  Else
   'We got enough free ram
   'Set last try to false
   LastTry = False
  End If
 End If
End Sub

Private Sub FreeMem_Click()
'Free Memory
Dim MemDat As MemoryStatus
Dim Tmp As String

On Local Error Resume Next

'Get the Memory Data
MemDat.dwLength = Len(MemDat)
GlobalMemoryStatus MemDat
'Show the Stop Icon in tray
Pic.Picture = IL1.ListImages(1).ExtractIcon
UpdateIcon
'The whole Trick
'Reserve all ram for a String
'then delete the string
Tmp = String(CLng(CLng(MemDat.dwTotalPhys) / CLng(2)), " ")
Tmp = vbNullChar
End Sub

