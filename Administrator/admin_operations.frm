VERSION 5.00
Begin VB.Form admin_operations 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "VTube - Administrator Operations"
   ClientHeight    =   7935
   ClientLeft      =   300
   ClientTop       =   2250
   ClientWidth     =   15525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "admin_operations.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "Generate Data Report 3"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13560
      TabIndex        =   17
      Top             =   9120
      Width           =   4335
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Generate Data Report 2"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13560
      TabIndex        =   16
      Top             =   8160
      Width           =   4335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Generate Data Report 1"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13560
      TabIndex        =   15
      Top             =   7200
      Width           =   4335
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   600
      Top             =   3840
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   17520
      Top             =   720
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Delete Videos"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13560
      TabIndex        =   5
      Top             =   5280
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add Videos"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   13560
      TabIndex        =   4
      Top             =   4320
      Width           =   4335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000012&
      Caption         =   "Administrator Operations"
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
      Height          =   7455
      Left            =   12240
      TabIndex        =   2
      Top             =   2760
      Width           =   6975
      Begin VB.CommandButton Command4 
         Caption         =   "View Source Code"
         BeginProperty Font 
            Name            =   "Zrnic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         TabIndex        =   11
         Top             =   3480
         Width           =   4335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Delete Users"
         BeginProperty Font 
            Name            =   "Zrnic"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1320
         TabIndex        =   3
         Top             =   600
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Statistics"
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
      Height          =   7455
      Left            =   1200
      TabIndex        =   1
      Top             =   2760
      Width           =   7695
      Begin VB.Label Label9 
         BackColor       =   &H80000012&
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
         Height          =   855
         Left            =   3960
         TabIndex        =   14
         Top             =   6120
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000012&
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
         Height          =   735
         Left            =   3960
         TabIndex        =   13
         Top             =   4560
         Width           =   1335
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000012&
         Caption         =   "Time"
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
         Height          =   855
         Left            =   4560
         TabIndex        =   12
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "No. Of Videos:"
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
         Height          =   735
         Left            =   1200
         TabIndex        =   9
         Top             =   6120
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "No. Of Users:"
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
         Height          =   615
         Left            =   1200
         TabIndex        =   8
         Top             =   4560
         Width           =   2175
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         Caption         =   "Version 4.19 (u60)"
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
         Height          =   735
         Left            =   1320
         TabIndex        =   7
         Top             =   2880
         Width           =   4815
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         Caption         =   "Date"
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
         Height          =   615
         Left            =   1680
         TabIndex        =   6
         Top             =   1080
         Width           =   1815
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
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
      Height          =   375
      Left            =   18120
      TabIndex        =   10
      Top             =   720
      Width           =   1695
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
      Height          =   735
      Left            =   7680
      TabIndex        =   0
      Top             =   1080
      Width           =   5775
   End
End
Attribute VB_Name = "admin_operations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private cn As ADODB.Connection
Dim str As String
Dim str1 As String
Dim str2 As String
Dim cn1 As ADODB.Connection
Dim cn2 As ADODB.Connection
Private Sub Command1_Click()
Unload Me
admin_user_screen.Show
End Sub

Private Sub Command2_Click()
Unload Me
admin_add_video.Show
End Sub

Private Sub Command3_Click()
Unload Me
admin_video_delete.Show
End Sub

Private Sub Command4_Click()
Call ShellExecute(0, "Open", "C:\Documents and Settings\Gaurav\My Documents\VTube - Video Database\Complete.vbp", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub Command5_Click()
DataReport1.Show
End Sub

Private Sub Command6_Click()
DataReport2.Show
End Sub

Private Sub Command7_Click()
DataReport4.Show
End Sub

Private Sub Form_Load()
Dim str1 As Integer
Dim str2 As String
Label2.Caption = DateValue(Now)
Label7.Caption = TimeValue(Now)
Set cn = New ADODB.Connection
str = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn.Open str
Dim rs As ADODB.Recordset

Dim qy1 As New ADODB.Command
With qy1
.CommandText = "get_user_id"
.CommandType = adCmdStoredProc
.ActiveConnection = cn
.Parameters.Append .CreateParameter _
(, adNumeric, adParamOutput)
.Parameters.Append .CreateParameter _
(, adNumeric, adParamOutput)
End With
qy1.Execute
Label8.Caption = qy1(0)
Label9.Caption = qy1(1)
End Sub

Private Sub Label6_Click()
Label6.ForeColor = &HFF&
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Unload Me
Load admin_login
admin_login.Show
End Sub

Private Sub Timer2_Timer()
Label7.Caption = TimeValue(Now)
End Sub
