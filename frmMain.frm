VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "MSDXM.OCX"
Object = "{DD8772A8-B496-11D2-8502-EF41B5E6386B}#1.0#0"; "JKJOYSTICK2.OCX"
Begin VB.Form pcjuke 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   8880
   ClientLeft      =   1215
   ClientTop       =   630
   ClientWidth     =   11880
   ControlBox      =   0   'False
   FillColor       =   &H00E0E0E0&
   ForeColor       =   &H00E0E0E0&
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   8880
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Height          =   5805
      Left            =   285
      Picture         =   "frmMain.frx":21FC1
      ScaleHeight     =   5700
      ScaleMode       =   0  'User
      ScaleWidth      =   5160
      TabIndex        =   29
      Top             =   2625
      Width           =   5220
   End
   Begin VB.PictureBox Picture1 
      Height          =   5765
      Left            =   285
      Picture         =   "frmMain.frx":2D104
      ScaleHeight     =   5700
      ScaleWidth      =   5160
      TabIndex        =   28
      Top             =   2625
      Width           =   5220
   End
   Begin VB.OptionButton hireOption 
      Caption         =   "Hire Off"
      Height          =   270
      Index           =   1
      Left            =   4050
      TabIndex        =   25
      Top             =   2325
      Value           =   -1  'True
      Width           =   960
   End
   Begin VB.OptionButton hireOption 
      Caption         =   "Hire On"
      Height          =   285
      Index           =   0
      Left            =   4050
      TabIndex        =   24
      Top             =   2025
      Width           =   945
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "d:\pcjuke\juke.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   288
      Left            =   4884
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "songs"
      Top             =   8856
      Visible         =   0   'False
      Width           =   984
   End
   Begin VB.CommandButton dirdwn 
      Caption         =   "dir down"
      Height          =   624
      Left            =   2415
      TabIndex        =   21
      Top             =   2025
      Width           =   804
   End
   Begin VB.CommandButton dir 
      Caption         =   "dir up"
      Height          =   624
      Left            =   1620
      TabIndex        =   20
      Top             =   2025
      Width           =   804
   End
   Begin VB.ListBox m3uList 
      Height          =   255
      Left            =   7110
      TabIndex        =   18
      Top             =   8865
      Visible         =   0   'False
      Width           =   3516
   End
   Begin JKJoystick2.JKJoystick JKJoystick1 
      Left            =   585
      Top             =   9225
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.CommandButton exitbut 
      Caption         =   "exit"
      Height          =   624
      Left            =   3210
      TabIndex        =   16
      Top             =   2025
      Width           =   804
   End
   Begin VB.ListBox playlist 
      Height          =   255
      ItemData        =   "frmMain.frx":3837C
      Left            =   2280
      List            =   "frmMain.frx":38383
      TabIndex        =   13
      Top             =   8640
      Visible         =   0   'False
      Width           =   2910
   End
   Begin VB.CommandButton add 
      Caption         =   "add"
      Height          =   624
      Left            =   825
      TabIndex        =   9
      Top             =   2025
      Width           =   804
   End
   Begin VB.CommandButton credit 
      Caption         =   "credit"
      Height          =   624
      Left            =   36
      TabIndex        =   8
      Top             =   2025
      Width           =   804
   End
   Begin VB.Timer Timer 
      Interval        =   1
      Left            =   60
      Top             =   9225
   End
   Begin MSComctlLib.ListView listview1 
      Height          =   8310
      Left            =   6015
      TabIndex        =   19
      Top             =   510
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   14658
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File Name"
         Object.Width           =   10460
      EndProperty
   End
   Begin VB.Label dubLabel 
      BackColor       =   &H00000000&
      Caption         =   """song selected once already, please select another"""
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   30
      Top             =   1560
      Width           =   5775
   End
   Begin Project1.VolumeControl VolumeWave 
      Left            =   4845
      Top             =   9210
      _ExtentX        =   847
      _ExtentY        =   847
      Volume          =   100
   End
   Begin Project1.VolumeControl VolumeMaster 
      Left            =   5775
      Top             =   9195
      _ExtentX        =   847
      _ExtentY        =   847
      Volume          =   90
      DeviceToControl =   0
   End
   Begin Project1.VolumeControl Volumeline 
      Left            =   6540
      Top             =   9210
      _ExtentX        =   847
      _ExtentY        =   847
      Volume          =   90
      DeviceToControl =   1
   End
   Begin VB.Label dubsong 
      Height          =   285
      Left            =   1320
      TabIndex        =   27
      Top             =   8880
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Last Selection"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   45
      TabIndex        =   26
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label datelab 
      Height          =   375
      Left            =   120
      TabIndex        =   22
      Top             =   8925
      Width           =   1575
   End
   Begin VB.Label dirtitle 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6045
      TabIndex        =   17
      Top             =   105
      Width           =   5865
   End
   Begin MediaPlayerCtl.MediaPlayer Player 
      Height          =   90
      Left            =   360
      TabIndex        =   14
      Top             =   8880
      Visible         =   0   'False
      Width           =   525
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
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
      PlayCount       =   1
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
      Volume          =   -380
      WindowlessVideo =   0   'False
   End
   Begin VB.Label joystat 
      BackColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   5055
      TabIndex        =   12
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label cointotal 
      BackColor       =   &H00C0C0C0&
      Height          =   300
      Left            =   5055
      TabIndex        =   11
      Top             =   2355
      Width           =   855
   End
   Begin VB.Label lastselbox 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   435
      Left            =   1620
      TabIndex        =   10
      Top             =   1560
      Width           =   4320
   End
   Begin VB.Label credtitleLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00000006&
      Caption         =   "credits"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   492
      Left            =   4920
      TabIndex        =   6
      Top             =   540
      Width           =   1015
   End
   Begin VB.Label creditLabel 
      Alignment       =   2  'Center
      BackColor       =   &H00000006&
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   492
      Left            =   4920
      TabIndex        =   5
      Top             =   1020
      Width           =   1015
   End
   Begin VB.Label songname 
      Alignment       =   2  'Center
      BackColor       =   &H00000006&
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   468
      Left            =   1452
      TabIndex        =   3
      Top             =   540
      Width           =   3492
   End
   Begin VB.Label Artistname 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000006&
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   528
      Left            =   1416
      TabIndex        =   4
      Top             =   996
      Width           =   3552
   End
   Begin VB.Label clock 
      Alignment       =   2  'Center
      BackColor       =   &H00000006&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "H:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   3081
         SubFormatType   =   4
      EndProperty
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   405
      Left            =   40
      TabIndex        =   2
      Top             =   1110
      Width           =   1410
   End
   Begin VB.Label nowplay 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      Caption         =   "Now Playing"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   510
      Left            =   40
      TabIndex        =   1
      Top             =   585
      Width           =   1425
   End
   Begin VB.Label hidenow 
      BackColor       =   &H00000000&
      Height          =   660
      Left            =   40
      TabIndex        =   15
      Top             =   555
      Width           =   1620
   End
   Begin VB.Label Addmore 
      Alignment       =   2  'Center
      BackColor       =   &H00000004&
      Caption         =   "insert $1 more for 3 plays"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   510
      Left            =   40
      TabIndex        =   7
      Top             =   45
      Width           =   5895
   End
   Begin VB.Label insertcoin 
      Alignment       =   2  'Center
      BackColor       =   &H80000006&
      Caption         =   "Insert $1 coin for play"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   510
      Left            =   40
      TabIndex        =   0
      Top             =   45
      Width           =   5895
   End
   Begin VB.Label hired 
      Alignment       =   2  'Center
      BackColor       =   &H00000004&
      Caption         =   "please select a song - no coins needed"
      BeginProperty Font 
         Name            =   "LcdD"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   510
      Left            =   40
      TabIndex        =   23
      Top             =   45
      Width           =   5895
   End
End
Attribute VB_Name = "pcjuke"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cointot As String
Dim afile As String
Dim coin As String
Dim lop As String
Dim dfile As String
Dim lopo As String
Dim sfile As String
Dim snext As String
Dim password As String
Dim FileName As String
Dim counto As String
Dim counta As String
Dim ListArray(1 To 9) As String
Dim FileOpen As Boolean
Dim i As Long
Dim Totalcredit As Long
Dim retval As Long
Dim creditadd As Long
Dim tinseconden As Long
Dim lengths, lengths1, min, sec As Long
Dim filenum As Integer
Dim intArrayNo As Integer
Dim CurrentTag As TagInfo


Dim FilePath$, thedata$, tmpString$, tmpstring2$, FindComma%

 Dim db As Database
    Dim rs As Recordset

Private Type TagInfo
    Tag As String * 3
    songname As String * 30
    artist As String * 30
    End Type
    
    Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
         ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Sub add_Click()
'service screen button

Call songadd
End Sub


Private Sub credtitleLabel_Click()
'display service screen (god mode)
cointotal.Caption = cointot

If credit.Visible = False Then
joystat.Visible = True
exitbut.Visible = True
credit.Visible = True
cointotal.Visible = True
add.Visible = True
dir.Visible = True
dirdwn.Visible = True
hireOption(0).Visible = True
hireOption(1).Visible = True

Else
joystat.Visible = False
exitbut.Visible = False
credit.Visible = False
cointotal.Visible = False
add.Visible = False
dir.Visible = False
dirdwn.Visible = False
hireOption(0).Visible = False
hireOption(1).Visible = False

End If
End Sub

Private Sub dir_Click()
'service screen button

Call dirchangeup 'change directory
End Sub

Private Sub exitbut_Click()
'service screen button

Unload pcjuke 'exits program
End Sub

Private Sub dirdwn_Click()
'service screen button

Call dirchangedown 'change directory
End Sub


Private Sub Form_Unload(Cancel As Integer)

datelab.Caption = Date & " " & Time ' sets a time stamp for database

With rs ' adds power off time to database
         .AddNew
                !songname = "power off"
                !artist = "unit shutdown correctly"
                !credit = cointotal
                !Date = datelab.Caption
                !hired = hireOption(0).Value
                
         .Update
      End With



End Sub

Private Sub hireOption_Click(Index As Integer)
 If hireOption(1) = True Then 'if not in hired mode show coin screen
        insertcoin.Visible = True
         Picture1.Visible = True
          Picture2.Visible = False
  End If

 If hireOption(0) = True Then 'if in hired mode hide coin screen
        insertcoin.Visible = False
        Picture2.Visible = True
          Picture1.Visible = False
  End If

End Sub

Private Sub hire()

 If insertcoin.Visible = False Then 'if not in hired mode show coin screen
        insertcoin.Visible = True
         Picture1.Visible = True
          Picture2.Visible = False
          hireOption(1) = True
          hireOption(0) = False
          GoTo endo
  End If

 If insertcoin.Visible = True Then 'if in hired mode hide coin screen
        insertcoin.Visible = False
        Picture2.Visible = True
          Picture1.Visible = False
          hireOption(1) = False
          hireOption(0) = True
  End If
endo:
End Sub

Private Sub jkJoystick1_ButtonPressed(ID As Integer)
' gets joystick button value

 Select Case ID
 
 Case 1
 joystat.Caption = "select"
 Call songadd
 
 Case 2
 joystat.Caption = "credit"
 Call creditbut
 
 Case 3
 joystat.Caption = "hire"
 Call hire
 
 Case 4
 joystat.Caption = "shutdown"
 Call ShutDown
 
 End Select
ends:
End Sub

Private Sub upstream()
' 38 is the character code for the up key
              
         Call keybd_event(38, 0, 0, 0)

         lop = 0
Stat: 'slow joy responce
         lopo = lop + 1
         lop = lopo
         If lopo = 9000 Then
         Exit Sub
         Else: GoTo Stat
        End If
End Sub

Private Sub downstream()
' 40 is the character code for the down key
        
         Call keybd_event(40, 0, 0, 0)
            lop = 0
Stat: 'slow joy responce
         lopo = lop + 1
         lop = lopo
         If lopo = 9000 Then
         Exit Sub
         Else: GoTo Stat
        End If
End Sub

Private Sub dirchangeup() 'change playlist file up

Starto:
    On Error GoTo ArrayErr
                 
   counta = counto + 1
   
   counto = counta

     intArrayNo = counta
    

  
Call listload


Call keybd_event(38, 0, 0, 0) ' set focus on listview
    
    If counta = 1 Then
    dirtitle.Caption = "Recent Chart"
    End If
    
    If counta = 2 Then
    dirtitle.Caption = "Hits of the 90s"
    End If
    
    If counta = 3 Then
    dirtitle.Caption = "Hits of the 80s"
    End If
    
    If counta = 4 Then
    dirtitle.Caption = "Hits of the 70s"
    End If
    
    If counta = 5 Then
    dirtitle.Caption = "Aussie Rock"
    End If
    
    If counta = 6 Then
    dirtitle.Caption = "Oldies"
    End If
    
    If counta = 7 Then
    dirtitle.Caption = "Disco"
    End If
    
    If counta = 8 Then
    dirtitle.Caption = "Classics"
    End If
    
     If counta = 9 Then
    dirtitle.Caption = "Party"
    counto = 0
    End If
    
    
    
    
    Exit Sub
    
ArrayErr:

    If Err.Number = 9 Then
       counto = 0
       GoTo Starto
    End If
    
End Sub
Private Sub dirchangedown()  'change playlist file down

Starto:
   On Error GoTo ArrayErr
    
    If counto = 0 Then
    counta = 10
    counto = 10
    End If
                 
                 
   counta = counto - 1
   
   counto = counta

   
   
   
     intArrayNo = counta
  
Call listload


Call keybd_event(38, 0, 0, 0) ' set focus on listview
    
   
    
   If counta = 1 Then
    dirtitle.Caption = "Recent Chart"
    End If
    
    If counta = 2 Then
    dirtitle.Caption = "Hits of the 90s"
    End If
    
    If counta = 3 Then
    dirtitle.Caption = "Hits of the 80s"
    End If
    
    If counta = 4 Then
    dirtitle.Caption = "Hits of the 70s"
    End If
    
    If counta = 5 Then
    dirtitle.Caption = "Aussie Rock"
    End If
    
    If counta = 6 Then
    dirtitle.Caption = "Oldies"
    End If
    
    If counta = 7 Then
    dirtitle.Caption = "Disco"
    End If
    
    If counta = 8 Then
    dirtitle.Caption = "Classics"
    End If
    
     If counta = 9 Then
    dirtitle.Caption = "Party"
    End If
    
    
    Exit Sub
    
ArrayErr:

    If Err.Number = 9 Then
       counto = 0
       GoTo Starto
    End If
    
End Sub


Private Sub creditbut()
'coin slot credit only, adds 1 to credit log total
coin = cointot + 1
cointot = coin
cointotal.Caption = coin
Open App.Path & "\creditlog.txt" For Output As #1

 Print #1, coin
          
      Close #1

Call creditit

End Sub
Private Sub credit_click()
'god mode credit button, doesnt add to credit log total

password = InputBox("enter password")
If password <> "gerbil" Then
MsgBox "password is invalid"
GoTo endo
Else
Call creditit
endo:
End If

End Sub

Private Sub songadd()
Dim Mouse As New CMouse
Mouse.Y = CLng(CDbl(Mouse.Y + 2000))
Mouse.X = CLng(CDbl(Mouse.X + 2000))

dubLabel.Visible = False

datelab.Caption = Date & " " & Time ' sets a time stamp for database

start:
Dim xfile As String

      Addmore.Visible = False 'hide add more coin screen
               
         If hireOption(1) = True Then 'if not in hired mode jumps to credit check
         hireOption(0) = False
          
      GoTo norm
  End If
         
               
               
      If hireOption(0) = True Then ' if in hired mode skips credit check
     
      GoTo hireskip
  End If
  
     
norm:
     If creditLabel.Caption = 0 Then
   GoTo ends 'if no credits exit sub
End If

  
    
    Totalcredit = creditLabel - 1 'subtracts a credit
  creditLabel = Totalcredit
hireskip:
  
    sfile = "d:" & listview1.SelectedItem.Tag
  ' open the mp3 to get the artist info
   filenum = FreeFile
   On Error GoTo skip
   Open sfile For Binary As #filenum
With CurrentTag
 Get #filenum, FileLen(sfile) - 127, .Tag
  If Not .Tag = "TAG" Then
        lastselbox.Caption = "no artist info"
        Close #filenum
        GoTo notag
    End If
    Get #filenum, , .songname
   Get #filenum, , .artist
   
    Close #filenum
    
    dubsong.Caption = RTrim(.songname)
    
    
    If dubsong = lastselbox.Caption Then
    GoTo dsong
    End If
    
    lastselbox.Caption = RTrim(.songname)
    afile = RTrim(.artist)
  playlist.AddItem sfile ' add selected file to play list
  
      
      'if song playing skip next setion
     If nowplay.Visible = True Then
     GoTo endo
     End If
 
 
 'play selected file

notag:
 If nowplay.Visible = True Then
     GoTo endo
     End If
Player.FileName = sfile
Player.Play
nowplay.Visible = True

     Volumeline.Volume = 0
    
   'update the label boxes
    songname.Caption = RTrim(.songname)
    Artistname.Caption = RTrim(.artist)
End With
  'sfile = ""
  
  
  GoTo endo
  
  
  
skip:
Close #1
  lastselbox.Caption = "ERROR, unable to play selection, try another" 'if file missing give back credit
  afile = sfile
  
  If hireOption(0) = True Then ' if in hired mode skips credit check
      insertcoin.Visible = False
      GoTo endo
  End If
   Totalcredit = creditLabel + 1
  creditLabel = Totalcredit
  
  GoTo endo
  
dsong:
  dubLabel.Visible = True
  If hireOption(0) = True Then ' if in hired mode skips credit check
      insertcoin.Visible = False
      GoTo endo
  End If
   Totalcredit = creditLabel + 1
  creditLabel = Totalcredit
  
  GoTo ends
  
  
endo:
      With rs ' adds song info to database
         .AddNew
                !songname = lastselbox.Caption
                !artist = afile
                !credit = cointotal
                !Date = datelab.Caption
                !hired = hireOption(0).Value
                
         .Update
      End With
ends:

End Sub

Private Sub creditit()

'detirmines how many coins inserted for bonus credit

If Addmore.Visible = False Then GoTo ad 'check the add more coin screen

Totalcredit = creditadd + 2 'add a credit to current level
creditadd = Totalcredit 'reset current level
creditLabel = Totalcredit 'display new level
Addmore.Visible = False 'hide add more coin screen

GoTo ends

ad:
creditadd = creditLabel.Caption

Totalcredit = creditadd + 1 'add a credit to current level
Addmore.Visible = True 'unhide add more coin screen
creditadd = Totalcredit 'reset current level
creditLabel = Totalcredit 'display new level
GoTo ends

ends:
End Sub

Private Sub Form_Load()
  Picture1.Visible = True
   Picture2.Visible = False
   hireOption(1) = True
pcjuke.Visible = True
'reset artist names
sfile = ""
  afile = ""
   songname.Caption = ""
    Artistname.Caption = ""

Volumeline.Volume = 70

'ShockwaveFlash.c
counto = 0
    ' load playlists into an array
    ListArray(1) = (App.Path & "\recent.m3u")
    ListArray(2) = (App.Path & "\90s.m3u")
    ListArray(3) = (App.Path & "\80s.m3u")
    ListArray(4) = (App.Path & "\70s.m3u")
    ListArray(5) = (App.Path & "\ausrock.m3u")
    ListArray(6) = (App.Path & "\oldies.m3u")
    ListArray(7) = (App.Path & "\disco.m3u")
    ListArray(8) = (App.Path & "\classics.m3u")
    ListArray(9) = (App.Path & "\party.m3u")
 Call dirchangeup
        



If JKJoystick1.CorrectlyConnected = False Then 'check if joystick plugged in
joystat.Caption = "no joystick"
GoTo skip
End If

JKJoystick1.TStart ' start the joystick control
skip:
nowplay.Visible = False

clock.Caption = ""



Call keybd_event(40, 0, 0, 0) ' set focus on filelist
Call keybd_event(40, 0, 0, 0) ' set focus on filelist


' hide service screen (god mode)
joystat.Visible = False
exitbut.Visible = False
credit.Visible = False
cointotal.Visible = False
add.Visible = False
dir.Visible = False
dirdwn.Visible = False
datelab.Visible = False
hireOption(0).Visible = False
hireOption(1).Visible = False
dubLabel.Visible = False

 Addmore.Visible = False 'hide add more coin screen
 
On Error GoTo Skip2
  
Open (App.Path & "\creditlog.txt") For Input As #1 ' get previous coin total


Input #1, cointot


Close #1

GoTo endo

Skip2: ' write a new file if creditlog.txt is missing
Open (App.Path & "\creditlog.txt") For Output As #1
'Print #1, counto
Input #1, cointot

Close #1

cointot = 0

endo:
creditadd = 0 'set the credit level to 0
creditLabel = creditadd 'display the credit level

 Set db = OpenDatabase(App.Path & "\juke.mdb") ' opens database
    Set rs = db.OpenRecordset("songs", dbOpenDynaset)

datelab.Caption = Date & " " & Time ' sets a time stamp for database

 With rs ' adds power on time to database
         .AddNew
                !songname = "power on"
                !artist = "unit started correctly"
                !credit = cointot
                !Date = datelab.Caption
                !hired = " "
                
         .Update
End With

End Sub

Private Sub JKJoystick1_PosChange(NewX As Long, NewY As Long, NewThrottle As Long, NewRudder As Long)

 'gets the stick value for left right
 
Select Case NewX
 Case Is < 5000
 joystat.Caption = "Left"
  lop = 0
Stat: 'slow joy responce
         lopo = lop + 1
         lop = lopo
         If lopo = 9000 Then
         Call dirchangeup
         Exit Sub
         Else: GoTo Stat
        End If
 
 Case Is > 60000
 joystat.Caption = "right"
 lop = 0
Stat2: 'slow joy responce
         lopo = lop + 1
         lop = lopo
         If lopo = 9000 Then
         Call dirchangedown
         Exit Sub
         Else: GoTo Stat2
        End If
 
End Select
 
 'gets the stick value for up down
 
 Select Case NewY
 Case Is < 5000
 joystat.Caption = "Up"
Call upstream
 
 Case Is > 60000
 joystat.Caption = "Down"
 Call downstream
 End Select
 

End Sub



Private Sub Player_EndOfStream(ByVal Result As Long)

nowplay.Visible = False
 
   On Error GoTo soo
    
    'goto next item on list and play
    last = last + 1
    snext = playlist.List(last)
    playlist.ListIndex = playlist.ListIndex + 1
 Player.FileName = snext
Player.Play

nowplay.Visible = True

soo:
last = playlist.ListIndex
' clear the label boxes in case of last song
songname.Caption = ""
Artistname.Caption = ""
clock.Caption = ""

 
 'remove top item from list
   If playlist.ListIndex = -1 Then
    Else
    playlist.RemoveItem playlist.ListIndex
    
    End If
 'if no list exit sub
   If playlist.ListCount = 0 Then
   
   Volumeline.Volume = 70
   
   
   Exit Sub
   
   End If

    
     'open the mp3 to get the artist info
 Open snext For Binary As #1
With CurrentTag
  Get #1, FileLen(snext) - 127, .Tag
    If Not .Tag = "TAG" Then
        songname.Caption = "No Artist Information"
          Artistname.Caption = ""
        Close #1
        Exit Sub
    End If
    Get #1, , .songname
    Get #1, , .artist
   
    Close #1

    
    'update the label boxes
    songname.Caption = RTrim(.songname)
    Artistname.Caption = RTrim(.artist)
    
      End With

 

End Sub




Private Sub Timer_Timer()


'used for time remaining
lengths = Player.Duration
tinseconden = Player.CurrentPosition
lengths1 = lengths - tinseconden
min = lengths1 \ 60
sec = lengths1 - min * 60
clock.Caption = min & " : " & sec


End Sub
Private Sub listload()
'gets the info from m3u playlist file

m3uList.Clear
listview1.ListItems.Clear
Dim i%

Open ListArray(intArrayNo) For Input As #1

Do
Line Input #1, thedata$
m3uList.AddItem thedata$
Loop While Not (EOF(1))
Close #1

For i% = 1 To m3uList.ListCount - 1
    If Left(m3uList.List(i%), 7) = "#EXTINF" Then 'see if a title was left
        FindComma% = InStr(1, m3uList.List(i%), ",")
        tmpString$ = Right(m3uList.List(i%), Len(m3uList.List(i%)) - FindComma%)
        If InStr(1, m3uList.List(i% + 1), "\") = 0 Then
            tmpstring2$ = FilePath$ & m3uList.List(i% + 1)
        Else
            tmpstring2$ = m3uList.List(i% + 1)
        End If
        listview1.ListItems.add , , tmpString$
        listview1.ListItems.item(listview1.ListItems.Count).Tag = tmpstring2$
        i% = i% + 1
    Else 'no title, use filename
       listview1.ListItems.add , , m3uList.List(i%)
        listview1.ListItems.item(listview1.ListItems.Count).Tag = FilePath$ & m3uList.List(i%)
    End If
Next

End Sub


