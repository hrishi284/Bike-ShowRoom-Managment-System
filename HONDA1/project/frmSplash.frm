VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   10065
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   16650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10065
   ScaleWidth      =   16650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2800
      Left            =   165
      ScaleHeight     =   2805
      ScaleWidth      =   3015
      TabIndex        =   2
      Top             =   700
      Width           =   3015
      Begin VB.PictureBox Picture3 
         Height          =   4815
         Left            =   -150
         ScaleHeight     =   4755
         ScaleWidth      =   4635
         TabIndex        =   3
         Top             =   -600
         Width           =   4695
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   4000
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
      Begin VB.PictureBox WebBrowser1 
         Height          =   4815
         Left            =   -150
         ScaleHeight     =   4755
         ScaleWidth      =   4635
         TabIndex        =   1
         Top             =   -360
         Width           =   4695
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   120
      Top             =   240
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
      TabIndex        =   4
      Top             =   0
      Width           =   3495
   End
   Begin VB.Image Image1 
      Height          =   11520
      Left            =   0
      Picture         =   "frmSplash.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   17595
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
   MDIForm1.Show
End Sub


Private Sub Frame1_Click()
Unload Me
MDIForm1.Show
End Sub

