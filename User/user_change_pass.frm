VERSION 5.00
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form user_change_pass 
   BackColor       =   &H00000000&
   Caption         =   "VTube - Change Passowrd"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin glxpbuttonz.UserButtonz Command1 
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   2400
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Done"
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
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1440
      Width           =   2895
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
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   1560
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Confirm New Password:"
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
      Height          =   1095
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Enter New Password:"
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
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "user_change_pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim string1 As Integer
Dim qy As ADODB.Command
Private cn As ADODB.Connection
Dim str As String
Dim str1 As String
Dim str2 As String
Dim cn1 As ADODB.Connection
Dim cn2 As ADODB.Connection
Dim obj1 As ListItem
Dim ctr As Integer

Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "Enter New Password", vbOKOnly + vbExclamation, "New Password"
Text1.SetFocus
ElseIf (Text1.Text <> Text2.Text) Then
MsgBox "Password Confirmation Incorrect", vbOKOnly + vbExclamation, "Wrong Cofirmation"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
Else
MsgBox "The password has been changed", vbOKOnly + vbInformation, "Password Change"
'query to change password in main database
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
str1 = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn.Open str1
Dim qy1 As ADODB.Command
str = "update userscreen set password = ? where id= ? "
Set qy1 = New ADODB.Command
With qy1
.CommandText = str
.CommandType = adCmdText
.ActiveConnection = cn
.Parameters.Append .CreateParameter _
(, adVarChar, adParamInput, 15)
.Parameters.Append .CreateParameter _
(, adNumeric, adParamInput)
End With

qy1(0) = Text1.Text
qy1(1) = Val(user_select_video.Label6.Caption)
Set rs = qy1.Execute

Set rs = Nothing
Set cn = Nothing

Unload Me
user_select_video.Enabled = True
Load user_select_video
user_select_video.Show
user_select_video.Label10.ForeColor = &HFFFFFF
End If
End Sub

