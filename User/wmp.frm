VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{E3583FCE-0595-4681-9ACD-48F7805DEFE1}#1.0#0"; "glxpbuttonz.ocx"
Begin VB.Form user_play_video 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "VTube - Play Video"
   ClientHeight    =   11010
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   19215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "wmp.frx":0000
   ScaleHeight     =   11010
   ScaleWidth      =   19215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   8040
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   5
      Left            =   5160
      Top             =   1920
   End
   Begin glxpbuttonz.UserButtonz Command2 
      Height          =   735
      Left            =   960
      TabIndex        =   18
      ToolTipText     =   "Click to go back to video select screen"
      Top             =   10080
      Width           =   3615
      _ExtentX        =   6376
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
      Caption         =   "Back To Video Select"
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
      Interval        =   500
      Left            =   19680
      Top             =   7320
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   19560
      Top             =   840
   End
   Begin VB.TextBox Text1 
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
      ForeColor       =   &H80000005&
      Height          =   1095
      Left            =   6240
      TabIndex        =   6
      Top             =   9720
      Width           =   13095
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5040
      Top             =   9240
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
            Picture         =   "wmp.frx":20CB
            Key             =   "video"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6975
      Left            =   240
      TabIndex        =   5
      ToolTipText     =   "Double click on any video to play it"
      Top             =   2280
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   12303
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
      Picture         =   "wmp.frx":2400
   End
   Begin VB.Label Label16 
      BackColor       =   &H80000007&
      Caption         =   "seconds"
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
      Left            =   17040
      TabIndex        =   19
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
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
      Left            =   16080
      TabIndex        =   17
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label14 
      BackColor       =   &H80000012&
      Caption         =   "Run Time:"
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
      Left            =   15000
      TabIndex        =   16
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label13 
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
      Height          =   375
      Left            =   3360
      TabIndex        =   15
      Top             =   9480
      Width           =   975
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      Caption         =   "Videos In Playlist:"
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
      Left            =   1080
      TabIndex        =   14
      Top             =   9480
      Width           =   2295
   End
   Begin VB.Label Label11 
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
      Left            =   12600
      TabIndex        =   13
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000012&
      Caption         =   "Video Department:"
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
      Left            =   10320
      TabIndex        =   12
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label9 
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
      Left            =   8640
      TabIndex        =   11
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000012&
      Caption         =   "Video Year:"
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
      Left            =   7200
      TabIndex        =   10
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label7 
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
      Left            =   18720
      TabIndex        =   9
      ToolTipText     =   "Log out of VTube"
      Top             =   240
      Width           =   1575
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
      Left            =   16440
      TabIndex        =   8
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
      Caption         =   "Joined On:"
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
      Left            =   15000
      TabIndex        =   7
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label4 
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
      Left            =   11880
      TabIndex        =   4
      Top             =   600
      Width           =   2655
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "User Name:"
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
      Left            =   10320
      TabIndex        =   3
      Top             =   600
      Width           =   1455
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
      Left            =   9000
      TabIndex        =   2
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000012&
      Caption         =   "Playlist No.:"
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
      Left            =   7200
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin WMPLibCtl.WindowsMediaPlayer wmp1 
      Height          =   7215
      Left            =   5880
      TabIndex        =   0
      Top             =   2160
      Width           =   13815
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   24368
      _cy             =   12726
   End
End
Attribute VB_Name = "user_play_video"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim b As WMPLibCtl.IWMPPlaylist
Dim c(100) As WMPLibCtl.IWMPMedia
Dim qy As ADODB.Command
Private cn As ADODB.Connection
Dim str As String
Dim str1 As String
Dim str2 As String
Dim str3 As String
Dim str4 As String
Dim cn1 As ADODB.Connection
Dim cn2 As ADODB.Connection
Dim obj1 As ListItem
Dim value As Integer
Dim dur As Integer
Dim temp As ListItem
Dim var_string As String
Dim con2 As ADODB.Connection

Private Sub Command2_Click()
Unload Me
user_select_video.Enabled = True
Load user_select_video
user_select_video.Show
End Sub

Private Sub Form_Load()
wmp1.settings.volume = 100
wmp1.windowlessVideo = True
Dim X As Integer
With ListView1
.ColumnHeaders.Clear
.ColumnHeaders.Add , , "Name", .Width * 0.3
.ColumnHeaders.Add , , "Path", .Width * 1.28
End With

Label2.Caption = user_select_video.Label2.Caption
Label4.Caption = user_select_video.Label5.Caption
Label6.Caption = user_select_video.Label8.Caption

Dim i As Integer
i = 0
Set cn = New ADODB.Connection
str = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
cn.Open str
Set rs = New ADODB.Recordset
str1 = "select distinct * from video v,playlist p where v.vid=p.vid and p.pno=? "

Set qy = New ADODB.Command
With qy
.CommandText = str1
.CommandType = adCmdText
.ActiveConnection = cn
.Parameters.Append .CreateParameter _
(, adNumeric, adParamInput)
End With
qy(0) = Val(Label2.Caption)

Set b = wmp1.newPlaylist(Label2.Caption, " ")
Set rs = qy.Execute

With rs

While Not .EOF
Set obj1 = ListView1.ListItems.Add(, , rs(0), , "video")
obj1.SubItems(1) = rs(4)
i = i + 1
str2 = rs(4)
Set c(i) = wmp1.newMedia(rs(4))
b.appendItem c(i)
.MoveNext
Wend
End With
Set rs = Nothing
Set cn = Nothing

wmp1.currentPlaylist = b
wmp1.Controls.play

End Sub


Private Sub Label7_Click()
Label7.ForeColor = &HFF&
Timer1.Enabled = True
End Sub

Private Sub ListView1_DblClick()
While (wmp1.currentMedia.sourceURL <> ListView1.SelectedItem.SubItems(1))
wmp1.Controls.Next
Wend
wmp1.Controls.play
End Sub

Private Sub Timer1_Timer()
Unload Me
user_select_video.Enabled = True
Unload user_select_video
user_sign_in.Show
End Sub

Private Sub Timer2_Timer()
Set con2 = New ADODB.Connection
Set rs = New ADODB.Recordset
Label13.Caption = wmp1.currentPlaylist.Count
Label15.Caption = wmp1.currentMedia.duration

Dim qy1 As ADODB.Command
Set qy1 = New ADODB.Command

str3 = "SERVER=vtube;" & "DRIVER={Microsoft ODBC for Oracle};" & "UID=system;PWD=manager;"
con2.Open str3
str4 = "SELECT * FROM video WHERE path=?"
With qy1
    .CommandText = str4
    .CommandType = adCmdText
    .ActiveConnection = con2
    .Parameters.Append .CreateParameter(, adVarChar, adParamInput, 200)
End With
var_string = Text2.Text
qy1(0) = var_string
Set rs = qy1.Execute

    Label9.Caption = rs(2)
    Label11.Caption = rs(3)
    Text1.Text = rs(5)

Set rs = Nothing
Set con2 = Nothing
'Timer2.Enabled = False
End Sub


Private Sub Timer3_Timer()
Text2.Text = wmp1.currentMedia.sourceURL
Timer3.Enabled = False
Timer2.Enabled = True
End Sub

Private Sub wmp1_OpenStateChange(ByVal NewState As Long)
Timer3.Enabled = True
End Sub
