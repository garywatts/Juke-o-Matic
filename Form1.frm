VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Volume Control"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   164
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   214
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Mute"
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   1560
      Width           =   855
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   7
      Top             =   1560
      Width           =   1500
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   1500
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1500
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1500
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1500
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1500
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1500
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      _Version        =   393216
      Max             =   100
      TickFrequency   =   10
   End
   Begin Project1.VolumeControl VolumeControl1 
      Left            =   2085
      Top             =   495
      _ExtentX        =   847
      _ExtentY        =   847
      Volume          =   66
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Check1_Click()
If Check1.Value = vbChecked Then
    VolumeControl1.Mute = True
Else
    VolumeControl1.Mute = False
End If

End Sub

Private Sub Form_Load()
Slider1.Value = VolumeControl1.Volume
Option1(0).Caption = "Master"
Option1(1).Caption = "LineIn"
Option1(2).Caption = "Microphone"
Option1(3).Caption = "Synthesizer"
Option1(4).Caption = "CD"
Option1(5).Caption = "Wave"
Option1(6).Caption = "Auxiliary"
Option1(VolumeControl1.DeviceToControl).Value = True

End Sub

Private Sub Option1_Click(Index As Integer)
VolumeControl1.DeviceToControl = Index
Dim vv As Long
vv = VolumeControl1.Volume
If VolumeControl1.Mute = True Then
    Check1.Value = vbChecked
Else
    Check1.Value = vbUnchecked
End If

End Sub

Private Sub Slider1_Scroll()
VolumeControl1.Volume = Slider1.Value

End Sub

Private Sub VolumeControl1_MuteChanged(NewMute As Boolean)
If NewMute Then
    Check1.Value = vbChecked
Else
    Check1.Value = vbUnchecked
End If


End Sub

Private Sub VolumeControl1_VolumeChanged(NewVolume As Long)
Slider1.Value = NewVolume

End Sub
