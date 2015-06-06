VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form admin_video_delete 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "VTube - Delete Video"
   ClientHeight    =   9690
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15720
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "admin_video_delete.frx":0000
   ScaleHeight     =   9690
   ScaleWidth      =   15720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3600
      Top             =   3120
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
            Picture         =   "admin_video_delete.frx":20CB
            Key             =   "video"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7455
      Left            =   5040
      TabIndex        =   5
      Top             =   3120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   13150
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
      ForeColor       =   16777215
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
      Picture         =   "admin_video_delete.frx":2400
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
      Left            =   720
      TabIndex        =   4
      Top             =   9960
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete Video"
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
      Height          =   735
      Left            =   15600
      TabIndex        =   2
      Top             =   6120
      Width           =   3015
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   17880
      Top             =   600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Video Deletion"
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
      Left            =   8160
      TabIndex        =   3
      Top             =   2160
      Width           =   3375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
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
      Left            =   6960
      TabIndex        =   1
      Top             =   1080
      Width           =   5775
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
      Height          =   495
      Left            =   18600
      TabIndex        =   0
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "admin_video_delete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim msgreturn As Integer
Dim obj1 As ListItem
Private cn As ADODB.Connection
Dim str As String
Dim str1 As String
Dim str2 As String
Dim cn1 As ADODB.Connection
Dim cn2 As ADODB.Connection
Private Sub Combo1_Click()
Combo2.Visible = True
End Sub

Private Sub Combo2_Click()
Combo3.Visible = True
End Sub

Private Sub Combo3_Click()
List1.Visible = True
Command1.Visible = True
End Sub

Private Sub Command1_Click()
msgreturn = MsgBox("Are you sure you want to delete this video?", 3 + vbQuestion, "Confirmation")
If msgreturn = 6 Then
MsgBox "Video Deleted", vbOKOnly + vbInformation, "Administrator"
'Procedure to delete video

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
.CommandText = "del_vid"
.CommandType = adCmdStoredProc
.ActiveConnection = cn1
.Parameters.Append .CreateParameter _
(, adNumeric, adParamInput)
End With
qy1(0) = id
qy1.Execute
ListView1.ListItems.Clear

Set cn1 = New ADODB.Connection
str = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn1.Open str
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "SELECT * from video", cn1
With rs
While Not .EOF
Set obj1 = ListView1.ListItems.Add(, , rs(0), , "video")
obj1.SubItems(1) = rs(1)
obj1.SubItems(2) = rs(2)
obj1.SubItems(3) = rs(3)
.MoveNext
Wend
End With
Set rs = Nothing
Set cn1 = Nothing

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
.ColumnHeaders.Add , , "Video_Id", .Width * 0.18
.ColumnHeaders.Add , , "Video Name", .Width * 0.55
.ColumnHeaders.Add , , "Year", .Width * 0.11
.ColumnHeaders.Add , , "Department", .Width * 0.15
.ColumnHeaders.Add , , "Path", .Width * 0.7
End With

ListView1.ListItems.Clear
Set cn1 = New ADODB.Connection
str = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn1.Open str
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
rs.Open "SELECT * from video", cn1
With rs
While Not .EOF
Set obj1 = ListView1.ListItems.Add(, , rs(0), , "video")
obj1.SubItems(1) = rs(1)
obj1.SubItems(2) = rs(2)
obj1.SubItems(3) = rs(3)
obj1.SubItems(4) = rs(4)
.MoveNext
Wend
End With
Set rs = Nothing
Set cn1 = Nothing

End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label6_Click()
Label6.ForeColor = &HFF&
Timer1.Enabled = True
End Sub

Private Sub Listview1_Click()
Command1.Enabled = True
End Sub

Private Sub ListView1_DblClick()
admin_video_delete.Enabled = False
admin_play_video.Show
If Me.Enabled = False Then
Me.Hide
End If
End Sub

Private Sub Timer1_Timer()
Unload Me
Load admin_login
admin_login.Show
End Sub
