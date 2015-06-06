VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form start_screen 
   Appearance      =   0  'Flat
   BackColor       =   &H80000006&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   6690
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Enter"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4560
      TabIndex        =   2
      Top             =   5040
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   10680
      Top             =   4320
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   735
      Left            =   600
      TabIndex        =   1
      Top             =   3360
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   1296
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   720
      Top             =   1200
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   3000
      Picture         =   "start_screen.frx":0000
      ScaleHeight     =   2175
      ScaleWidth      =   5895
      TabIndex        =   0
      Top             =   480
      Width           =   5895
   End
End
Attribute VB_Name = "start_screen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Const WM_USER = &H400
Const CCM_FIRST = &H2000&
Const CCM_SETBKCOLOR = (CCM_FIRST + 1)
Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
Const PBM_SETBARCOLOR = (WM_USER + 9)
Dim tlevel As Integer

Private Sub Command1_Click()
Unload Me
main_form.Show
End Sub

Private Sub Form_Load()
ChangePBForeColour ProgressBar1.hWnd, vbWhite
ChangePBBackColour ProgressBar1.hWnd, vbBlack
tlevel = 5
TransForm Me, Val(tlevel)
End Sub

Private Sub Timer1_Timer()
tlevel = tlevel + 5
TransForm Me, Val(tlevel)
If tlevel = 255 Then
Timer1.Enabled = False
Timer2.Enabled = True
End If
End Sub

Private Sub Timer2_Timer()
If ProgressBar1.value <= 40 Then
ProgressBar1.value = ProgressBar1.value + 0.4
ElseIf ProgressBar1.value >= 40 And ProgressBar1.value <= 60 Then
ProgressBar1.value = ProgressBar1.value + 0.2
ElseIf ProgressBar1.value >= 60 And ProgressBar1.value <= 80 Then
ProgressBar1.value = ProgressBar1.value + 0.5
ElseIf ProgressBar1.value >= 80 And ProgressBar1.value <= 99.5 Then
ProgressBar1.value = ProgressBar1.value + 0.5
Else
ProgressBar1.value = 100
Command1.Visible = True
'MsgBox "Full"
Timer2.Enabled = False
End If
End Sub

Private Function ChangePBForeColour(ByVal hWnd As Long, ByVal lColor As Long)
    'Change colour of bar
    SendMessage hWnd, PBM_SETBARCOLOR, 0, ByVal lColor
End Function


Private Function ChangePBBackColour(ByVal hWnd As Long, ByVal lColor As Long)
    'Change colour of background
    SendMessage hWnd, PBM_SETBKCOLOR, 0, ByVal lColor
End Function
