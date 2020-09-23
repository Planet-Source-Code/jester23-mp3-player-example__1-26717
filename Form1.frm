VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "MP3 Player Example"
   ClientHeight    =   4800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   120
      TabIndex        =   30
      Top             =   5040
      Width           =   7095
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   9240
      Top             =   2760
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   8520
      Top             =   2760
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8040
      Top             =   2640
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   8040
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open MP3"
      Filter          =   "MP3 Files [*.mp3]|*.mp3"
   End
   Begin VB.CheckBox Check1 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6720
      TabIndex        =   22
      ToolTipText     =   "Repeat"
      Top             =   720
      Value           =   1  'Checked
      Width           =   495
   End
   Begin MediaPlayerCtl.MediaPlayer MP3 
      Height          =   255
      Left            =   7920
      TabIndex        =   21
      Top             =   3840
      Visible         =   0   'False
      Width           =   2175
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   -1  'True
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   0
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin MSComctlLib.ProgressBar PB 
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Playlist"
      Height          =   1935
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   7095
      Begin VB.ListBox PlayList 
         BackColor       =   &H00800000&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   1545
         IntegralHeight  =   0   'False
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "Right Click For Menu"
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.PictureBox Picture15 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   7050
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   15
      Top             =   150
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox Picture14 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   7050
      Picture         =   "Form1.frx":03DF
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   14
      Top             =   150
      Width           =   225
   End
   Begin VB.PictureBox Picture13 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   6720
      Picture         =   "Form1.frx":07BE
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   13
      Top             =   150
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.PictureBox Picture12 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   6720
      Picture         =   "Form1.frx":0B75
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   12
      Top             =   150
      Width           =   225
   End
   Begin VB.PictureBox PicBorder 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   6540
      TabIndex        =   0
      Top             =   720
      Width           =   6540
      Begin VB.PictureBox Picture11 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   5280
         Picture         =   "Form1.frx":0F2C
         ScaleHeight     =   300
         ScaleWidth      =   1200
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.PictureBox Picture10 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   5280
         Picture         =   "Form1.frx":1508
         ScaleHeight     =   300
         ScaleWidth      =   1200
         TabIndex        =   9
         Top             =   0
         Width           =   1200
      End
      Begin VB.PictureBox Picture9 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   3960
         Picture         =   "Form1.frx":1AE4
         ScaleHeight     =   300
         ScaleWidth      =   1200
         TabIndex        =   8
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.PictureBox Picture8 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   3960
         Picture         =   "Form1.frx":2017
         ScaleHeight     =   300
         ScaleWidth      =   1200
         TabIndex        =   7
         Top             =   0
         Width           =   1200
      End
      Begin VB.PictureBox Picture7 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   2640
         Picture         =   "Form1.frx":254A
         ScaleHeight     =   300
         ScaleWidth      =   1200
         TabIndex        =   6
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.PictureBox Picture6 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   2640
         Picture         =   "Form1.frx":2AAF
         ScaleHeight     =   300
         ScaleWidth      =   1200
         TabIndex        =   5
         Top             =   0
         Width           =   1200
      End
      Begin VB.PictureBox Picture5 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   1320
         Picture         =   "Form1.frx":3014
         ScaleHeight     =   300
         ScaleWidth      =   1200
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   1320
         Picture         =   "Form1.frx":3521
         ScaleHeight     =   300
         ScaleWidth      =   1200
         TabIndex        =   3
         Top             =   0
         Width           =   1200
      End
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   0
         Picture         =   "Form1.frx":3A2E
         ScaleHeight     =   300
         ScaleWidth      =   1200
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   0
         Picture         =   "Form1.frx":3F65
         ScaleHeight     =   300
         ScaleWidth      =   1200
         TabIndex        =   1
         Top             =   0
         Width           =   1200
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5880
      TabIndex        =   29
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   7920
      TabIndex        =   28
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   8160
      TabIndex        =   27
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   8520
      TabIndex        =   26
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   8760
      TabIndex        =   25
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   9120
      TabIndex        =   24
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   9360
      TabIndex        =   23
      Top             =   3360
      Width           =   255
   End
   Begin VB.Label lblAuthor 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   1560
      Width           =   7095
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1200
      Width           =   7095
   End
   Begin VB.Shape Shape1 
      Height          =   4695
      Left            =   0
      Top             =   0
      Width           =   7335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MP3  Player Example"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Click here to drag"
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SYSCOMMAND = &H112

Private Sub Label1_MouseDown(Button As Integer, shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        ReleaseCapture
        SendMessage Me.hwnd, WM_SYSCOMMAND, &HF012, 0
    End If
End Sub

Private Sub PlayList_MouseUp(Button As Integer, shift As Integer, x As Single, y As Single)
If Button = 2 Then
    Form2.PopupMenu Form2.mnupopup
Else
    DoEvents
End If
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
MP3.PlayCount = 1
End If

If Check1.Value = 2 Then
MP3.PlayCount = 0
End If
End Sub

Private Sub Picture11_Click()
If PlayList.Enabled = True Then
PlayList.Enabled = False
Else
PlayList.Enabled = False
PlayList.Enabled = True
End If
End Sub

Private Sub Picture13_Click()
Form1.WindowState = vbMinimized
End Sub

Private Sub Picture15_Click()
MP3.Stop
End
End Sub

Private Sub Form_Load()
Me.Width = Shape1.Width + 5
Me.Height = Shape1.Height + 5
End Sub

Private Sub Picture2_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Picture2.Visible = False
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = False
Picture6.Visible = True
Picture7.Visible = False
Picture8.Visible = True
Picture9.Visible = False
Picture10.Visible = True
Picture11.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Picture2.Visible = True
Picture3.Visible = False
Picture4.Visible = True
Picture5.Visible = False
Picture6.Visible = True
Picture7.Visible = False
Picture8.Visible = True
Picture9.Visible = False
Picture10.Visible = True
Picture11.Visible = False
Picture12.Visible = True
Picture13.Visible = False
Picture14.Visible = True
Picture15.Visible = False
End Sub

Private Sub Picture3_Click()
On Error Resume Next
CD.filename = MP3.filename
CD.ShowOpen
Label4.Caption = 0
Label5.Caption = 0
Label6.Caption = 0
Label7.Caption = 0
Label8.Caption = 0
Label9.Caption = 0
If CD.filename = MP3.filename Then
GoTo Crapy
Else
MP3.filename = CD.filename
PB.Max = MP3.Duration
Timer1.Enabled = True
Timer2.Enabled = True
Timer3.Enabled = True
lblTitle.Caption = MP3.GetMediaInfoString(mpClipTitle)
lblAuthor.Caption = MP3.GetMediaInfoString(mpClipAuthor)
PlayList.AddItem List1.ListCount + 1 & " " & MP3.GetMediaInfoString(mpClipTitle)
List1.AddItem MP3.filename
End If

Crapy:
End Sub

Private Sub Picture4_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Picture2.Visible = True
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = False
Picture8.Visible = True
Picture9.Visible = False
Picture10.Visible = True
Picture11.Visible = False
End Sub

Private Sub Picture5_Click()
MP3.Play
Timer3.Enabled = True
End Sub

Private Sub Picture6_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Picture2.Visible = True
Picture3.Visible = False
Picture4.Visible = True
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = True
Picture8.Visible = True
Picture9.Visible = False
Picture10.Visible = True
Picture11.Visible = False
End Sub

Private Sub Picture7_Click()
Timer3.Enabled = False
MP3.Pause
End Sub

Private Sub Picture8_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Picture2.Visible = True
Picture3.Visible = False
Picture4.Visible = True
Picture5.Visible = False
Picture6.Visible = True
Picture7.Visible = False
Picture8.Visible = False
Picture9.Visible = True
Picture10.Visible = True
Picture11.Visible = False
End Sub

Private Sub Picture10_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Picture2.Visible = True
Picture3.Visible = False
Picture4.Visible = True
Picture5.Visible = False
Picture6.Visible = True
Picture7.Visible = False
Picture8.Visible = True
Picture9.Visible = False
Picture10.Visible = False
Picture11.Visible = True
End Sub

Private Sub Picture12_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Picture12.Visible = False
Picture13.Visible = True
Picture14.Visible = True
Picture15.Visible = False
End Sub

Private Sub Picture14_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Picture12.Visible = True
Picture13.Visible = False
Picture14.Visible = False
Picture15.Visible = True
End Sub

Private Sub PicBorder_MouseMove(Button As Integer, shift As Integer, x As Single, y As Single)
Picture2.Visible = True
Picture3.Visible = False
Picture4.Visible = True
Picture5.Visible = False
Picture6.Visible = True
Picture7.Visible = False
Picture8.Visible = True
Picture9.Visible = False
Picture10.Visible = True
Picture11.Visible = False
End Sub

Private Sub Picture9_Click()
CD.filename = ""
MP3.Stop
lblTitle.Caption = ""
lblAuthor.Caption = ""
PB.Value = 0
Timer3.Enabled = False
Label4.Caption = 0
Label5.Caption = 0
Label6.Caption = 0
Label7.Caption = 0
Label8.Caption = 0
Label9.Caption = 0
End Sub

Private Sub PlayList_Click()
List1.ListIndex = PlayList.ListIndex
MP3.filename = List1.Text
Label4.Caption = 0
Label5.Caption = 0
Label6.Caption = 0
Label7.Caption = 0
Label8.Caption = 0
Label9.Caption = 0
Timer3.Enabled = True
End Sub

Private Sub Timer1_Timer()
PB.Value = MP3.CurrentPosition
End Sub

Private Sub Timer2_Timer()
Label9.Caption = Label9.Caption + 1
If Label9.Caption = 10 Then
Label9.Caption = 0
Label8.Caption = Label8.Caption + 1
End If
If Label8.Caption = 6 Then
Label8.Caption = 0
Label7.Caption = Label7.Caption + 1
End If
If Label7.Caption = 10 Then
Label7.Caption = 0
Label6.Caption = Label6.Caption + 1
End If
If Label6.Caption = 6 Then
Label6.Caption = 0
Label5.Caption = Label5.Caption + 1
End If
If Label5.Caption = 10 Then
Label5.Caption = 0
Label4.Caption = Label4.Caption + 1
End If
End Sub

Private Sub Timer3_Timer()
Label2.Caption = Label4.Caption & Label5.Caption & " : " & Label6.Caption & Label7.Caption & " : " & Label8.Caption & Label9.Caption
End Sub

Private Sub MP3_EndOfStream(ByVal Result As Long)
Label4.Caption = 0
Label5.Caption = 0
Label6.Caption = 0
Label7.Caption = 0
Label8.Caption = 0
Label9.Caption = 0
End Sub
