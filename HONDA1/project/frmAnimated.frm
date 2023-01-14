VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmAnimated 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3495
   FillColor       =   &H00004000&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "frmAnimated.frx":0000
   ScaleHeight     =   4485
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   120
      Top             =   240
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2800
      Left            =   165
      ScaleHeight     =   2805
      ScaleWidth      =   3015
      TabIndex        =   0
      Top             =   700
      Width           =   3015
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   4815
         Left            =   -150
         TabIndex        =   1
         Top             =   -600
         Width           =   4695
         ExtentX         =   8281
         ExtentY         =   8493
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   135
      Left            =   1080
      TabIndex        =   3
      Top             =   3840
      Visible         =   0   'False
      Width           =   375
      URL             =   "D:\project\bikesound.mp3"
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
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   661
      _cy             =   238
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BackStyle       =   0  'Transparent
      Caption         =   "Please wait while your request is processing"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmAnimated"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
WebBrowser1.Navigate App.Path & "\loading.gif"
End Sub

Private Sub Timer1_Timer()
Unload Me
MDIForm1.Show



End Sub
