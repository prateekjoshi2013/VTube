VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form admin_user_screen 
   BackColor       =   &H80000007&
   BorderStyle     =   0  'None
   Caption         =   "VTube - User Screening"
   ClientHeight    =   9600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19305
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "admin_user_screen.frx":0000
   ScaleHeight     =   9600
   ScaleWidth      =   19305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "admin_user_screen.frx":20CB
            Key             =   "user"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6015
      Left            =   5280
      TabIndex        =   5
      Top             =   3000
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   10610
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483643
      BackColor       =   -2147483642
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Zrnic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
      Picture         =   "admin_user_screen.frx":291A
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
      Left            =   600
      TabIndex        =   3
      Top             =   9840
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete User"
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
      Left            =   8640
      TabIndex        =   2
      Top             =   9840
      Width           =   3495
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   17280
      Top             =   960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "User Screening"
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
      Height          =   495
      Left            =   7560
      TabIndex        =   4
      Top             =   2040
      Width           =   5415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
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
      Height          =   615
      Left            =   7320
      TabIndex        =   1
      Top             =   960
      Width           =   5895
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
      Left            =   18240
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "admin_user_screen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msgreturn As Integer
Private cn As ADODB.Connection
Dim str As String
Dim str1 As String
Dim str2 As String
Dim cn1 As ADODB.Connection
Dim cn2 As ADODB.Connection
Dim obj1 As ListItem
Private Sub Command1_Click()
msgreturn = MsgBox("Are you sure you want to ban this user?", 3 + vbQuestion, "Confirmation")
If msgreturn = 6 Then
MsgBox "User Removed from Database", vbOKOnly + vbInformation, "Ban User"

'Procedure to delete user
Dim id As Integer
Dim id2 As String

id2 = ListView1.SelectedItem.Text
id = CInt(id2)
Set cn1 = New ADODB.Connection
str1 = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn1.Open str1
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
Dim qy1 As New ADODB.Command
With qy1
.CommandText = "del_user"
.CommandType = adCmdStoredProc
.ActiveConnection = cn1
.Parameters.Append .CreateParameter _
(, adNumeric, adParamInput)
End With
qy1(0) = id
qy1.Execute
ListView1.ListItems.Clear
Set cn = New ADODB.Connection
str = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn.Open str
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "SELECT * from userscreen", cn
With rs
While Not .EOF
Set obj1 = ListView1.ListItems.Add(, , rs(1), , "user")
obj1.SubItems(1) = rs(0)
obj1.SubItems(2) = rs(4)
obj1.SubItems(3) = rs(5)
.MoveNext
Wend
End With
Set rs = Nothing
Set cn = Nothing


End If
End Sub

Private Sub Command2_Click()
Unload Me
Load admin_login
admin_operations.Show
End Sub

Private Sub Form_Load()
With ListView1
.ColumnHeaders.Clear
.ColumnHeaders.Add , , "User_Id", .Width * 0.13
.ColumnHeaders.Add , , "User Name", .Width * 0.35
.ColumnHeaders.Add , , "Joining Date", .Width * 0.2
.ColumnHeaders.Add , , "Contact Number", .Width * 0.31
ListView1.ListItems.Clear

Set cn = New ADODB.Connection
str = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn.Open str
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "SELECT * from userscreen", cn
With rs
While Not .EOF
Set obj1 = ListView1.ListItems.Add(, , rs(1), , "user")
obj1.SubItems(1) = rs(0)
obj1.SubItems(2) = rs(4)
obj1.SubItems(3) = rs(5)
.MoveNext
Wend
End With
Set rs = Nothing
Set cn = Nothing

End With
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label6_Click()
Label6.ForeColor = &HFF&
Timer1.Enabled = True
End Sub

Private Sub List1_Click()
'Will display table of users with reqd characteristics
Command1.Visible = True
End Sub

Private Sub Timer1_Timer()
Unload Me
Load admin_login
admin_login.Show
'exit to login screen
End Sub
