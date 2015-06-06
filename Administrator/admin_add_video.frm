VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form admin_add_video 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "VTube - Add Video"
   ClientHeight    =   7125
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19260
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "admin_add_video.frx":0000
   ScaleHeight     =   7125
   ScaleWidth      =   19260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Browse"
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
      Left            =   15360
      TabIndex        =   14
      Top             =   6240
      Width           =   2535
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   17880
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back to Operations"
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
      Left            =   1080
      TabIndex        =   13
      Top             =   9720
      Width           =   2895
   End
   Begin VB.ComboBox Combo5 
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
      Height          =   420
      Left            =   6120
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   4920
      Width           =   1335
   End
   Begin VB.ComboBox Combo4 
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
      Height          =   420
      Left            =   8280
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Video"
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
      Height          =   615
      Left            =   8400
      TabIndex        =   9
      Top             =   9720
      Width           =   2895
   End
   Begin VB.TextBox Text5 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   6120
      TabIndex        =   8
      Top             =   7680
      Width           =   8535
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000006&
      Enabled         =   0   'False
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
      Height          =   615
      Left            =   6120
      TabIndex        =   7
      Top             =   6240
      Width           =   8535
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      Enabled         =   0   'False
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
      Height          =   615
      Left            =   6120
      TabIndex        =   6
      Top             =   3480
      Width           =   8535
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   17880
      Top             =   600
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Adding Video"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6840
      TabIndex        =   10
      Top             =   2040
      Width           =   6015
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Description:"
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
      Left            =   2040
      TabIndex        =   5
      Top             =   8040
      Width           =   3255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Path:"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   6360
      Width           =   3255
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Video Specifications: "
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
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Video Name:"
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
      Left            =   2040
      TabIndex        =   2
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Welcome Administrator"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   840
      Width           =   5775
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Logout"
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
      Height          =   495
      Left            =   18600
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "admin_add_video"
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
Dim sub_count As Integer
Dim msgreturn As Integer

Private Sub Combo5_Click()
Combo4.Visible = True
Combo4.SetFocus
End Sub

Private Sub Command1_Click()
If Text1.Text = "" Then
MsgBox "You must choose a Video", vbInformation + vbOKOnly, "Information Missing"
ElseIf Text4.Text = "" Then
MsgBox "You must choose a Video", vbInformation + vbOKOnly, "Information Missing"
ElseIf Combo5.Text = "" Then
MsgBox "Select Year", vbOKOnly + vbInformation, "Information Missing"
Combo5.SetFocus
ElseIf Combo4.Text = "" Then
MsgBox "Select Department", vbOKOnly + vbInformation, "Information Missing"
ElseIf Len(Text5.Text) > 200 Then
MsgBox "Description must be less than 200 characters", vbOKOnly + vbInformation, "Invalid Information"
Else
msgreturn = MsgBox("Add this video to the database?", 3 + vbQuestion, "Confirmation")
If msgreturn = 6 Then
MsgBox "Video Entered", vbOKOnly + vbInformation, "Video Addition Successful"


Set cn1 = New ADODB.Connection
str1 = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn1.Open str1
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
Dim qy As New ADODB.Command
With qy
.CommandText = "add_vid"
.CommandType = adCmdStoredProc
.ActiveConnection = cn1
.Parameters.Append .CreateParameter _
(, adVarChar, adParamInput, 50)
.Parameters.Append .CreateParameter _
(, adVarChar, adParamInput, 5)
.Parameters.Append .CreateParameter _
(, adVarChar, adParamInput, 15)
.Parameters.Append .CreateParameter _
(, adVarChar, adParamInput, 100)
.Parameters.Append .CreateParameter _
(, adVarChar, adParamInput, 200)
.ActiveConnection = cn1
End With
qy(0) = Text1.Text
qy(1) = Combo5.Text
qy(2) = Combo4.Text
qy(3) = Text4.Text
qy(4) = Text5.Text
qy.Execute

Unload Me
admin_add_video.Show
End If
End If

End Sub

Private Sub Command2_Click()
Unload Me
Load admin_login
admin_operations.Show
End Sub

Private Sub Command3_Click()
Dim sFile As String
Dim title As String
With CommonDialog1
    .InitDir = "C:\VTubeVideos"
    .DialogTitle = "Add Video To Database"
    .CancelError = False
    .Filter = "All Suported Files"
    .ShowOpen
    If Len(.FileName) = 0 Then
        Exit Sub
    End If
    sFile = .FileName
    title = .FileTitle
With Text4
    .Text = sFile
End With
Text1.Text = title
End With
End Sub

Private Sub Form_Load()
Combo5.AddItem "F.E."
Combo5.AddItem "S.E."
Combo5.AddItem "T.E."
Combo5.AddItem "B.E."
Combo5.AddItem "-NA-"
Combo4.AddItem "Computer"
Combo4.AddItem "Mechanical"
Combo4.AddItem "I.T"
Combo4.AddItem "E&TC"
Combo4.AddItem "Electronics"
Combo4.AddItem "Civil"
Combo4.AddItem "-NA-"
End Sub

Private Sub Label6_Click()
Label6.ForeColor = &HFF&
Timer1.Enabled = True
End Sub

Private Sub Label9_Click()
Unload Me
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
Command1.Enabled = True
End Sub


Private Sub Timer1_Timer()
Unload Me
Load admin_login
admin_login.Show
End Sub
