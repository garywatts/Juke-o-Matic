VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Timer Timer1 
         Interval        =   3500
         Left            =   3870
         Top             =   1920
      End
      Begin VB.Label Licence 
         Caption         =   "Licenced for use at The Max Hotel Geelong"
         Height          =   225
         Left            =   3600
         TabIndex        =   7
         Top             =   3435
         Width           =   3165
      End
      Begin VB.Label label1 
         Caption         =   "Jukebox"
         BeginProperty Font 
            Name            =   "Futura XBlkIt BT"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   372
         Left            =   5112
         TabIndex        =   6
         Top             =   1368
         Width           =   1920
      End
      Begin VB.Image imgLogo 
         Height          =   3120
         Left            =   276
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   480
         Width           =   3048
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Copyright"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   2
         Top             =   2985
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         Caption         =   "Intown Entertainment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4560
         TabIndex        =   1
         Top             =   3195
         Width           =   2415
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1.4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6525
         TabIndex        =   3
         Top             =   2700
         Width           =   330
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5712
         TabIndex        =   4
         Top             =   2340
         Width           =   1140
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Juke o matic"
         BeginProperty Font 
            Name            =   "Futura XBlkIt BT"
            Size            =   26.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   645
         Left            =   3450
         TabIndex        =   5
         Top             =   810
         Width           =   3465
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'Dim retval As Long

Private Sub Form_Load()
'retval = auxSetVolume(0, &HFFFF)
'retval = auxSetVolume(13, &HFFFF)
pcjuke.VolumeMaster.Volume = 90
pcjuke.VolumeWave.Volume = 100
pcjuke.Volumeline.Volume = 90



pcjuke.Player.FileName = "d:\pcjuke\robot.mp3"
pcjuke.Player.Play

End Sub

Private Sub Timer1_Timer()

Load pcjuke
Unload frmSplash

End Sub

