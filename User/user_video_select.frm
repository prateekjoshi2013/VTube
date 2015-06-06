VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form user_select_video 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "VTube - Video Selection and Playlist Creation"
   ClientHeight    =   3090
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "user_video_select.frx":0000
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin glxpbuttonz.UserButtonz UserButtonz1 
      Height          =   615
      Left            =   16320
      TabIndex        =   17
      ToolTipText     =   "Deletes the selected item from your playlist"
      Top             =   9360
      Width           =   3015
      _ExtentX        =   5318
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
      Caption         =   "Delete"
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
      Left            =   10800
      TabIndex        =   15
      ToolTipText     =   "Click to transfer selected video to your playlist"
      Top             =   5760
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1296
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   ""
      IconHighLiteColor=   0
      CaptionHighLite =   -1  'True
      CaptionHighLiteColor=   255
      Picture         =   "user_video_select.frx":20CB
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
   Begin glxpbuttonz.UserButtonz Command3 
      Height          =   615
      Left            =   14880
      TabIndex        =   14
      ToolTipText     =   "Clears the entire playlist"
      Top             =   10080
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
      Caption         =   "Clear Playlist"
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
   Begin glxpbuttonz.UserButtonz Command2 
      Height          =   615
      Left            =   13200
      TabIndex        =   13
      ToolTipText     =   "Click to play your playlist"
      Top             =   9360
      Width           =   3015
      _ExtentX        =   5318
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
      Caption         =   "Play"
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
      Interval        =   500
      Left            =   19440
      Top             =   840
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   6120
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
            Picture         =   "user_video_select.frx":251D
            Key             =   "video"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00000000&
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
      ForeColor       =   &H00FFFFFF&
      Height          =   450
      Left            =   7920
      Style           =   2  'Dropdown List
      TabIndex        =   3
      ToolTipText     =   "Select Department"
      Top             =   1680
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00000000&
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
      Height          =   450
      Left            =   7920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Select Year"
      Top             =   480
      Width           =   3735
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   6015
      Left            =   12840
      TabIndex        =   1
      ToolTipText     =   "This is your playlist"
      Top             =   3120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   10610
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
      Picture         =   "user_video_select.frx":2852
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7575
      Left            =   1080
      TabIndex        =   0
      ToolTipText     =   "Select the video to add to your playlist"
      Top             =   2760
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   13361
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      PictureAlignment=   4
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   0
      BorderStyle     =   1
      Appearance      =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Zrnic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
      Picture         =   "user_video_select.frx":50BA
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "Department:"
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
      Left            =   6120
      TabIndex        =   19
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "Year:"
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
      Left            =   6480
      TabIndex        =   18
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Change Password"
      BeginProperty Font 
         Name            =   "Zrnic"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   14640
      TabIndex        =   16
      ToolTipText     =   "Click to change your password"
      Top             =   2520
      Width           =   3735
   End
   Begin VB.Label Label9 
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
      Height          =   615
      Left            =   18840
      TabIndex        =   12
      ToolTipText     =   "Log out of VTube"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
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
      Height          =   255
      Left            =   15600
      TabIndex        =   11
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000012&
      Caption         =   "Joined On: "
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
      Height          =   255
      Left            =   13080
      TabIndex        =   10
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
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
      Height          =   255
      Left            =   15600
      TabIndex        =   9
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
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
      Height          =   255
      Left            =   15600
      TabIndex        =   8
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      Caption         =   "UserID:"
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
      Height          =   255
      Left            =   13080
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "UserName:"
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
      Height          =   255
      Left            =   13080
      TabIndex        =   6
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000012&
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
      Height          =   255
      Left            =   15600
      TabIndex        =   5
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Playlist ID:"
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
      Height          =   255
      Left            =   13080
      TabIndex        =   4
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "user_select_video"
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

Private Sub Combo1_Click()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset

If ctr > 1 Then

ListView1.ListItems.Clear
str1 = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn.Open str1
Dim qy1 As ADODB.Command
str = "SELECT * FROM video WHERE year = ? "
str = str & " AND dept = ? "
Set qy1 = New ADODB.Command
With qy1
.CommandText = str
.CommandType = adCmdText
.ActiveConnection = cn
.Parameters.Append .CreateParameter _
(, adVarChar, adParamInput, 10)
.Parameters.Append .CreateParameter _
(, adVarChar, adParamInput, 15)
End With

qy1(0) = Combo1.Text
qy1(1) = Combo2.Text
Set rs = qy1.Execute
With rs
While Not .EOF
Set obj1 = ListView1.ListItems.Add(, , rs(0), , "video")
obj1.SubItems(1) = rs(1)
.MoveNext
Wend
End With


Set rs = Nothing
Set cn = Nothing

End If
Combo2.Enabled = True
End Sub

Private Sub Combo2_Click()
ctr = ctr + 2
ListView1.Enabled = True
Command1.Enabled = True
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset

ListView1.ListItems.Clear
str1 = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn.Open str1
Dim qy1 As ADODB.Command
str = "SELECT * FROM video WHERE year = ? "
str = str & " AND dept = ? "
Set qy1 = New ADODB.Command
With qy1
.CommandText = str
.CommandType = adCmdText
.ActiveConnection = cn
.Parameters.Append .CreateParameter _
(, adVarChar, adParamInput, 10)
.Parameters.Append .CreateParameter _
(, adVarChar, adParamInput, 15)
End With

qy1(0) = Combo1.Text
qy1(1) = Combo2.Text
Set rs = qy1.Execute
With rs
While Not .EOF
Set obj1 = ListView1.ListItems.Add(, , rs(0), , "video")
obj1.SubItems(1) = rs(1)
.MoveNext
Wend
End With


Set rs = Nothing
Set cn = Nothing
End Sub

Private Sub Command1_Click()
If ListView1.ListItems.Count = 0 Then
MsgBox "No Video Selected", vbOKCancel + vbInformation, "Select a Video"
Else
ListView2.ListItems.Clear
Set cn1 = New ADODB.Connection
str1 = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn1.Open str1
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
Dim qy As New ADODB.Command
With qy
.CommandText = "add_del_vid"
.CommandType = adCmdStoredProc
.ActiveConnection = cn1
.Parameters.Append .CreateParameter _
(, adNumeric, adParamInput)
.Parameters.Append .CreateParameter _
(, adVarChar, adParamInput, 50)
.Parameters.Append .CreateParameter _
(, adNumeric, adParamInput)
.ActiveConnection = cn1
End With

qy(0) = Val(ListView1.SelectedItem.Text)
qy(1) = ListView1.SelectedItem.SubItems(1)
qy(2) = Val(Label2.Caption)
qy.Execute
Set cn1 = Nothing
Set rs1 = Nothing

ListView1.ListItems.Remove (ListView1.SelectedItem.Index)

ListView2.ListItems.Clear
Set cn = New ADODB.Connection
str = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn.Open str
Set rs = New ADODB.Recordset
str1 = "select distinct vid ,vname,pno from playlist where pno= ? "

Set qy = New ADODB.Command
With qy
.CommandText = str1
.CommandType = adCmdText
.ActiveConnection = cn
.Parameters.Append .CreateParameter _
(, adNumeric, adParamInput)
End With
qy(0) = Val(Label2.Caption)

Set rs = qy.Execute
With rs
While Not .EOF
Set obj1 = ListView2.ListItems.Add(, , rs(0), , "video")
obj1.SubItems(1) = rs(1)
.MoveNext
Wend
End With
Set rs = Nothing
Set cn = Nothing

Set rs = Nothing
Set cn = Nothing
End If
End Sub

Private Sub Command2_Click()
If ListView2.ListItems.Count = 0 Then
MsgBox "Playlist is Empty!", vbOKOnly + vbExclamation, "Playlist Empty"
Else
Load user_play_video
user_play_video.Show
With Me
.Enabled = False
End With
If Me.Enabled = False Then
Me.Hide
End If
End If
'Unload Me
End Sub

Private Sub Command3_Click()
Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
'procedure to populate LABLES
str1 = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn.Open str1
Dim qy1 As ADODB.Command
str = "delete FROM playlist WHERE  pno=? "
Set qy1 = New ADODB.Command
With qy1
.CommandText = str
.CommandType = adCmdText
.ActiveConnection = cn
.Parameters.Append .CreateParameter _
(, adNumeric, adParamInput)
End With

qy1(0) = Val(Label2.Caption)

Set rs = qy1.Execute

Set rs = Nothing
Set cn = Nothing
ListView2.ListItems.Clear
End Sub

Private Sub Form_Load()
ctr = 0
Combo1.AddItem "F.E."
Combo1.AddItem "S.E."
Combo1.AddItem "T.E."
Combo1.AddItem "B.E."
Combo1.AddItem "-NA-"
Combo2.AddItem "Computer"
Combo2.AddItem "Mechanical"
Combo2.AddItem "I.T"
Combo2.AddItem "E&TC"
Combo2.AddItem "Electronics"
Combo2.AddItem "Civil"
Combo2.AddItem "-NA-"
With ListView1
.ColumnHeaders.Clear
.ColumnHeaders.Add , , "Video Number", .Width * 0.2
.ColumnHeaders.Add , , "Video Name", .Width * 0.79
End With
With ListView2
.ColumnHeaders.Clear
.ColumnHeaders.Add , , "Video Number", .Width * 0.35
.ColumnHeaders.Add , , "Video Name", .Width * 0.64
End With

Set cn = New ADODB.Connection
Set rs = New ADODB.Recordset
'procedure to populate LABELS
ListView1.ListItems.Clear
str1 = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn.Open str1
Dim qy1 As ADODB.Command
str = "SELECT * FROM userscreen WHERE id = ? "
Set qy1 = New ADODB.Command
With qy1
.CommandText = str
.CommandType = adCmdText
.ActiveConnection = cn
.Parameters.Append .CreateParameter _
(, adNumeric, adParamInput)
End With
string1 = Val(user_sign_in.Text1.Text)
qy1(0) = string1

Set rs = qy1.Execute

Label5.Caption = rs(0)
Label8.Caption = rs(4)
Label6.Caption = rs(1)
Label2.Caption = rs(3)
Set rs = Nothing
Set cn = Nothing

Set cn = New ADODB.Connection
str = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn.Open str
Set rs = New ADODB.Recordset
str1 = "select distinct vid ,vname,pno from playlist where pno= ? "

Set qy = New ADODB.Command
With qy
.CommandText = str1
.CommandType = adCmdText
.ActiveConnection = cn
.Parameters.Append .CreateParameter _
(, adNumeric, adParamInput)
End With
qy(0) = Val(Label2.Caption)

Set rs = qy.Execute
With rs
While Not .EOF
Set obj1 = ListView2.ListItems.Add(, , rs(0), , "video")
obj1.SubItems(1) = rs(1)
.MoveNext
Wend
End With
Set rs = Nothing
Set cn = Nothing

End Sub

Private Sub Label10_Click()
Label10.ForeColor = &HFF&
user_change_pass.Show
With Me
    .Enabled = False
End With
'MsgBox "Input Box for change password and confirmation"
End Sub

Private Sub Label9_Click()
Timer1.Enabled = True
Label9.ForeColor = &HFF&
End Sub

Private Sub Timer1_Timer()
Unload Me
user_sign_in.Show
End Sub

Private Sub UserButtonz1_Click()

Set cn = New ADODB.Connection
str = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn.Open str
Set rs = New ADODB.Recordset
str1 = "delete from playlist where vid = ? and pno = ? "

Set qy = New ADODB.Command
With qy
.CommandText = str1
.CommandType = adCmdText
.ActiveConnection = cn
.Parameters.Append .CreateParameter _
(, adNumeric, adParamInput)
.Parameters.Append .CreateParameter _
(, adNumeric, adParamInput)
End With

qy(0) = Val(ListView2.SelectedItem)
qy(1) = Val(Label2.Caption)

Set rs = qy.Execute
Set rs = Nothing
Set cn = Nothing

ListView2.ListItems.Remove (ListView2.SelectedItem.Index)
End Sub
