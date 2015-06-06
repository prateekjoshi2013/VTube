VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form user_sign_up 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "VTube - New User"
   ClientHeight    =   9510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15630
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "user_sign_up.frx":0000
   ScaleHeight     =   9510
   ScaleWidth      =   15630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin glxpbuttonz.UserButtonz Command2 
      Height          =   735
      Left            =   6840
      TabIndex        =   15
      ToolTipText     =   "Back to User login screen"
      Top             =   9840
      Width           =   4695
      _ExtentX        =   8281
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
      Caption         =   "Back to Login Screen"
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
      Height          =   615
      Left            =   7680
      TabIndex        =   14
      ToolTipText     =   "Enter your details and click to sign up"
      Top             =   6120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Zrnic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Sign Up!"
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
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   2520
      Top             =   7800
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   10680
      Top             =   6240
   End
   Begin VB.TextBox Text6 
      BackColor       =   &H80000006&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   735
      Left            =   9600
      TabIndex        =   11
      Top             =   8520
      Width           =   4695
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H80000006&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   735
      Left            =   9600
      TabIndex        =   10
      ToolTipText     =   "This is your User Id. Remember it!"
      Top             =   7200
      Width           =   4695
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   735
      Left            =   8760
      TabIndex        =   7
      ToolTipText     =   "Enter a contact number"
      Top             =   4800
      Width           =   5295
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   8760
      PasswordChar    =   "*"
      TabIndex        =   6
      ToolTipText     =   "Confirm your desireed password"
      Top             =   3600
      Width           =   5295
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   735
      IMEMode         =   3  'DISABLE
      Left            =   8760
      PasswordChar    =   "*"
      TabIndex        =   5
      ToolTipText     =   "Enter the desired password"
      Top             =   2400
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   735
      Left            =   8760
      TabIndex        =   4
      ToolTipText     =   "Enter your name"
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Login Id"
      BeginProperty Font 
         Name            =   "Starcraft"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   14760
      TabIndex        =   13
      Top             =   7440
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Important!"
      BeginProperty Font 
         Name            =   "Starcraft"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3120
      TabIndex        =   12
      Top             =   7440
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Your Playlist Id:"
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
      Left            =   5160
      TabIndex        =   9
      Top             =   8760
      Width           =   3855
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Your User Id:"
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
      Left            =   5160
      TabIndex        =   8
      Top             =   7440
      Width           =   3855
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Contact:"
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
      Left            =   5280
      TabIndex        =   3
      Top             =   5040
      Width           =   3015
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Confirm Password:"
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
      Left            =   5280
      TabIndex        =   2
      Top             =   3840
      Width           =   3015
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
      Left            =   5280
      TabIndex        =   1
      Top             =   2640
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "User Name: "
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
      Left            =   5880
      TabIndex        =   0
      Top             =   1440
      Width           =   1815
   End
End
Attribute VB_Name = "user_sign_up"
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

Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "User Name not entered", vbOKOnly + vbExclamation, "Invalid Information"
Text1.SetFocus
ElseIf (Text2.Text <> Text3.Text) Then
MsgBox "Password Confirmation Incorrect", vbOKCancel + vbExclamation, "Invalid Information"
Text2.Text = ""
Text3.Text = ""
Text2.SetFocus
ElseIf Text4.Text = "" Then
MsgBox "Contact Number not entered", vbOKOnly + vbExclamation, "Invalid Information"
Text4.SetFocus
ElseIf ((Len(Text4.Text) > 10) Or Len(Text4.Text) < 7) Then
MsgBox "Enter Valid Contact Number(7-10 numbers)", vbOKOnly + vbExclamation, "Invalid Information"
Text4.SetFocus
Else
Set cn1 = New ADODB.Connection
str1 = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn1.Open str1
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
Dim qy1 As New ADODB.Command
With qy1
.CommandText = "add_user"
.CommandType = adCmdStoredProc
.ActiveConnection = cn1
.Parameters.Append .CreateParameter _
(, adVarChar, adParamInput, 50)
.Parameters.Append .CreateParameter _
(, adVarChar, adParamInput, 50)
.Parameters.Append .CreateParameter _
(, adVarChar, adParamInput, 50)
.Parameters.Append .CreateParameter _
(, adNumeric, adParamOutput)
.Parameters.Append .CreateParameter _
(, adNumeric, adParamOutput)
End With

qy1(0) = Text1.Text
qy1(1) = Text2.Text
qy1(2) = Text4.Text
qy1.Execute
Text5.Text = qy1(3)
Text6.Text = qy1(4)
Timer1.Enabled = True
Timer2.Enabled = True

Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Text4.Enabled = False
Command1.Enabled = False
End If

End Sub

Private Sub Command2_Click()
Unload Me
user_sign_in.Show
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (((KeyAscii < 65) And (KeyAscii <> 32) And (KeyAscii <> 8)) Or ((KeyAscii > 91) And (KeyAscii < 97)) Or (KeyAscii > 122)) Then
MsgBox "Enter Alphabets Only", vbOKOnly + vbExclamation, "Invalid Information"
KeyAscii = 0
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (((KeyAscii < 48) And (KeyAscii <> 8)) Or (KeyAscii > 58)) Then
MsgBox "Enter Numbers Only", vbExclamation + vbOKOnly, "Invalid Information"
KeyAscii = 0
End If
End Sub

Private Sub Timer1_Timer()
Command2.Enabled = True
End Sub

Private Sub Timer2_Timer()
If Label7.Visible = True Then
Label7.Visible = False
Label8.Visible = False
Else
Label7.Visible = True
Label8.Visible = True
End If
End Sub
