VERSION 5.00
Begin VB.Form big_logo 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "Big logo"
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   ScaleHeight     =   4695
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   240
      Top             =   360
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3855
      Left            =   840
      Picture         =   "big_logo.frx":0000
      ScaleHeight     =   3855
      ScaleWidth      =   10575
      TabIndex        =   0
      Top             =   480
      Width           =   10575
      Begin VB.Timer Timer2 
         Interval        =   1000
         Left            =   9960
         Top             =   0
      End
   End
End
Attribute VB_Name = "big_logo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tlevel As Integer

Private Sub Form_Load()
tlevel = 5
TransForm Me, Val(tlevel)
End Sub

Private Sub Timer1_Timer()
tlevel = tlevel + 5
TransForm Me, Val(tlevel)
If tlevel = 255 Then
Timer1.Enabled = False
End If
End Sub

Private Sub Timer2_Timer()
Unload Me
start_screen.Show
End Sub
