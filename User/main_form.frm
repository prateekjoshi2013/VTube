VERSION 5.00
Begin VB.Form main_form 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "VTube - Welcome"
   ClientHeight    =   3090
   ClientLeft      =   300
   ClientTop       =   2250
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "main_form.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000012&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      Picture         =   "main_form.frx":300C42
      ScaleHeight     =   2055
      ScaleWidth      =   5895
      TabIndex        =   18
      Top             =   960
      Width           =   5895
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Caption         =   "Developers"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   720
      TabIndex        =   11
      Top             =   7560
      Width           =   4575
      Begin VB.Label Label13 
         BackColor       =   &H80000012&
         Caption         =   "Gaurav Phadke - 3330"
         BeginProperty Font 
            Name            =   "Zrnic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Top             =   2040
         Width           =   3855
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000012&
         Caption         =   "Mohammad Asad - 3361"
         BeginProperty Font 
            Name            =   "Zrnic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   1560
         Width           =   3855
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000012&
         Caption         =   "Sahil Magdum - 3359"
         BeginProperty Font 
            Name            =   "Zrnic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   1080
         Width           =   3855
      End
      Begin VB.Label Label10 
         BackColor       =   &H80000012&
         Caption         =   "Prateek Joshi -3362"
         BeginProperty Font 
            Name            =   "Zrnic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   3855
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4815
      Left            =   12480
      TabIndex        =   1
      Top             =   3120
      Width           =   7335
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         Caption         =   "User"
         BeginProperty Font 
            Name            =   "Unreal Tournament"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   1680
         TabIndex        =   3
         Top             =   3120
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "Administrator"
         BeginProperty Font 
            Name            =   "Unreal Tournament"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   1680
         TabIndex        =   2
         Top             =   960
         Width           =   5415
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7920
      TabIndex        =   0
      Top             =   9480
      Width           =   4455
   End
   Begin VB.Timer Timer1 
      Interval        =   15
      Left            =   9600
      Top             =   6600
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   " : A Database Management Mini Project"
      BeginProperty Font 
         Name            =   "Unreal Tournament"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   5760
      TabIndex        =   17
      Top             =   1680
      Width           =   9375
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000012&
      Caption         =   " - Developed with Microsoft Visual Basic 6.0 and Oracle9.0i"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   480
      TabIndex        =   16
      Top             =   6960
      Width           =   7815
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000007&
      Caption         =   " - Extremely stringent security measures for administrators"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   6480
      Width           =   7695
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000007&
      Caption         =   " - Detailed classification and sorting of videos with unparalleled        ease of access"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   480
      TabIndex        =   9
      Top             =   5640
      Width           =   7695
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000007&
      Caption         =   " - Supported on all versions of Microsoft Windows XP"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   5160
      Width           =   7695
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000007&
      Caption         =   " - Ultimate portability and versatility"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   4680
      Width           =   7695
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000007&
      Caption         =   " - Robust and failure-proof back-end database connectivity"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   4200
      Width           =   7695
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000007&
      Caption         =   " - Easy-to-use, user-friendly Graphical User Interface"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3720
      Width           =   7695
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000007&
      Caption         =   " - A comprehensive video database management software"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   3240
      Width           =   7695
   End
End
Attribute VB_Name = "main_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tlevel As Integer

Private Sub Command1_Click()
End
End Sub

Private Sub Form_Load()
tlevel = 15
TransForm Me, Val(tlevel)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.FontSize = 22
Label2.FontSize = 22
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.FontSize = 22
Label2.FontSize = 22
End Sub

Private Sub Label1_Click()
Unload Me
admin_login.Show
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.FontSize = 28
End Sub

Private Sub Label2_Click()
Unload Me
user_sign_in.Show
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.FontSize = 28
End Sub

Private Sub Timer1_Timer()
tlevel = tlevel + 15
TransForm Me, Val(tlevel)
If tlevel = 255 Then
Timer1.Enabled = False
End If
End Sub
