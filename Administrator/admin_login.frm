VERSION 5.00
Begin VB.Form admin_login 
   BackColor       =   &H80000008&
   BorderStyle     =   0  'None
   Caption         =   "VTube - Administrator Login"
   ClientHeight    =   9675
   ClientLeft      =   300
   ClientTop       =   2250
   ClientWidth     =   19245
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   15
      Left            =   3000
      Top             =   3120
   End
   Begin VB.Timer SecurityTimer 
      Interval        =   10000
      Left            =   1440
      Top             =   2640
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back To Main Screen"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9000
      TabIndex        =   12
      Top             =   9000
      Width           =   2895
   End
   Begin VB.PictureBox LogoBox 
      Appearance      =   0  'Flat
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   7440
      Picture         =   "admin_login.frx":0000
      ScaleHeight     =   2055
      ScaleWidth      =   5895
      TabIndex        =   10
      Top             =   840
      Width           =   5895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Administrator Login"
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
      Height          =   3615
      Left            =   6840
      TabIndex        =   4
      Top             =   5040
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton Command1 
         Caption         =   "Enter"
         BeginProperty Font 
            Name            =   "Zrnic"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2160
         TabIndex        =   11
         Top             =   2640
         Width           =   2895
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "Zrnic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   3120
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1800
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000006&
         BeginProperty Font 
            Name            =   "Zrnic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000005&
         Height          =   495
         Left            =   3120
         TabIndex        =   7
         Top             =   840
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "Password: "
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
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   1920
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "Login ID: "
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
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   3
      Left            =   11640
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   2
      Left            =   10680
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   1
      Left            =   9720
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   495
      Index           =   0
      Left            =   8760
      Shape           =   3  'Circle
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Welcome Administrator"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7080
      TabIndex        =   9
      Top             =   4200
      Width           =   6735
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Height          =   975
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   9960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Height          =   975
      Index           =   2
      Left            =   19440
      TabIndex        =   2
      Top             =   10080
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Height          =   975
      Index           =   1
      Left            =   19440
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Height          =   975
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "admin_login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim counter As Integer
Dim pass As String
Dim secure_flag As Boolean
Dim username As String
Dim tlevel As Integer

Private Sub Command1_Click()
If ((Text1.Text = "admin") And (Text2.Text = "admin")) Then
Unload Me
admin_operations.Show
ElseIf Text1.Text = "" Then
MsgBox "Please enter User Name", vbOKOnly + vbExclamation, "Access Denied"
ElseIf Text2.Text = "" Then
MsgBox "Please enter Password", vbOKOnly + vbExclamation, "Access Denied"
Else
MsgBox "Incorrect Username or Password", vbOKOnly + vbCritical, "Access Denied"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End If

End Sub

Private Sub Command2_Click()
'MsgBox "Exiting VTube", vbOKOnly + vbInformation, "VTube-Exit"
Unload Me
main_form.Show
End Sub

Private Sub Form_Load()
counter = 0
tlevel = 15
TransForm Me, Val(tlevel)
End Sub

Private Sub Label1_Click(Index As Integer)
Label1(Index).Enabled = False
Shape1(Index).Visible = True
counter = counter + Index
If counter = 6 Then
Label2.Visible = True
Label3.Visible = True
Text1.Visible = True
Text2.Visible = True
Frame1.Visible = True
LogoBox.Visible = True
Text1.SetFocus
secure_flag = True
End If
End Sub

Private Sub SecurityTimer_Timer()
If secure_flag = False Then
MsgBox "Unauthorised Access Attempted!", vbOKOnly + vbCritical, "Security Breach"
Unload Me
End If
End Sub

Private Sub Timer1_Timer()
tlevel = tlevel + 15
TransForm Me, Val(tlevel)
If tlevel = 255 Then
Timer1.Enabled = False
End If
End Sub
