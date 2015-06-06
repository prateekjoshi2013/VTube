VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form user_sign_in 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "VTube - User Sign In"
   ClientHeight    =   3090
   ClientLeft      =   240
   ClientTop       =   1800
   ClientWidth     =   4680
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   Picture         =   "user_sign_in.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   15
      Left            =   1080
      Top             =   3360
   End
   Begin glxpbuttonz.UserButtonz UserButtonz1 
      Height          =   735
      Left            =   8520
      TabIndex        =   8
      ToolTipText     =   "Back to Home Page"
      Top             =   8640
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Zrnic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Back to Main Screen"
      IconHighLiteColor=   0
      CaptionHighLite =   -1  'True
      CaptionHighLiteColor=   255
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   0
      ColorButtonUp   =   8421504
      ColorButtonDown =   0
      BorderBrightness=   0
      ColorBright     =   16777215
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin glxpbuttonz.UserButtonz Command1 
      Height          =   735
      Left            =   9000
      TabIndex        =   6
      ToolTipText     =   "Login to VTube"
      Top             =   6240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Zrnic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Sign In"
      IconHighLiteColor=   0
      CaptionHighLite =   -1  'True
      CaptionHighLiteColor=   255
      Style           =   1
      Checked         =   0   'False
      ColorButtonHover=   0
      ColorButtonUp   =   8421504
      ColorButtonDown =   0
      BorderBrightness=   0
      ColorBright     =   16777215
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   6000
      Top             =   7080
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   9120
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Enter your password"
      Top             =   5040
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   615
      Left            =   9120
      TabIndex        =   0
      ToolTipText     =   "Enter your User Id"
      Top             =   3720
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "User Login"
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
      Height          =   5295
      Left            =   6720
      TabIndex        =   2
      Top             =   2760
      Width           =   7575
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "New User? Sign Up!"
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
         Left            =   2280
         TabIndex        =   5
         ToolTipText     =   "Click to create your account"
         Top             =   4560
         Width           =   2895
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "Password:"
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
         Left            =   600
         TabIndex        =   4
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "User Id:"
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
         Left            =   600
         TabIndex        =   3
         Top             =   1080
         Width           =   1575
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   ": A Database Management Systems Mini-Project"
      BeginProperty Font 
         Name            =   "Unreal Tournament"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5880
      TabIndex        =   7
      Top             =   720
      Width           =   9615
   End
End
Attribute VB_Name = "user_sign_in"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private cn As ADODB.Connection
Dim str As String
Dim str1 As String
Dim str2 As String
Dim cn1 As ADODB.Connection
Dim cn2 As ADODB.Connection
Dim tlevel As Integer

Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Please Enter your User Id", vbOKOnly + vbInformation, "Access Denied"
Text1.SetFocus
ElseIf Text2.Text = "" Then
MsgBox "Please Enter your Password", vbOKOnly + vbInformation, "Access Denied"
Text2.SetFocus
Else
Set cn = New ADODB.Connection
str1 = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn.Open str1
Dim qy1 As ADODB.Command
str = "SELECT count(*) FROM userscreen WHERE id = ? "

Set qy1 = New ADODB.Command
With qy1
.CommandText = str
.CommandType = adCmdText
.ActiveConnection = cn
.Parameters.Append .CreateParameter _
(, adNumeric, adParamInput)
End With

qy1(0) = Val(Text1.Text)
Set rs = qy1.Execute

If (rs(0) = 0) Then
MsgBox "Enter valid Username and Password", vbOKOnly + vbCritical, "Access Denied"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus

Set rs = Nothing
Set cn = Nothing
Else
Set cn = New ADODB.Connection
str1 = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn.Open str1

str = "SELECT password FROM userscreen WHERE id = ? "

Set qy1 = New ADODB.Command
With qy1
.CommandText = str
.CommandType = adCmdText
.ActiveConnection = cn
.Parameters.Append .CreateParameter _
(, adNumeric, adParamInput)
End With

qy1(0) = Val(Text1.Text)
Set rs = qy1.Execute
If StrComp(rs(0), Text2.Text, vbTextCompare) = 0 Then
user_select_video.Show
Unload Me
Else
MsgBox "Please enter correct User Id and Password", vbOKOnly + vbCritical, "Access Denied"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End If
End If

Set rs = Nothing
Set cn = Nothing
End If
End Sub

Private Sub Form_Load()
tlevel = 15
TransForm Me, Val(tlevel)
End Sub

Private Sub Label3_Click()

'This should pop out when mouse comes over it!!

Label3.ForeColor = &HFF&
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Unload Me
user_sign_up.Show
End Sub

Private Sub Timer2_Timer()
tlevel = tlevel + 15
TransForm Me, Val(tlevel)
If tlevel = 255 Then
Timer2.Enabled = False
End If
End Sub

Private Sub UserButtonz1_Click()
Unload Me
main_form.Show
End Sub
