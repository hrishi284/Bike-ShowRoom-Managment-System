VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   8355
   ClientLeft      =   2790
   ClientTop       =   3105
   ClientWidth     =   12495
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4936.409
   ScaleMode       =   0  'User
   ScaleWidth      =   11732.13
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   3720
      MaskColor       =   &H00000000&
      Picture         =   "frmLogin.frx":30E8C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6000
      UseMaskColor    =   -1  'True
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2280
      Picture         =   "frmLogin.frx":313BE
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   6000
      Width           =   615
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   4320
      Width           =   1845
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   390
      Left            =   3360
      TabIndex        =   2
      Top             =   5040
      Width           =   1140
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   1680
      TabIndex        =   1
      Top             =   5040
      Width           =   1140
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   3120
      TabIndex        =   0
      Top             =   3480
      Width           =   1725
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Forgot Password"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   1
      Left            =   1560
      TabIndex        =   5
      Top             =   4320
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Index           =   0
      Left            =   1560
      TabIndex        =   4
      Top             =   3480
      Width           =   1920
   End
   Begin VB.Image Image1 
      Height          =   15375
      Left            =   -2280
      Picture         =   "frmLogin.frx":319D5
      Stretch         =   -1  'True
      Top             =   -3480
      Width           =   23730
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Private Sub cmdOK_Click()
'If txtID.Text <> "" And txtPassword.Text <> "" Then

'Module1.numval ("select count(*)from login where  ID ='" & txtID.Text & "' and password='" & txtPassword.Text & "'")

'If module.nm = 1 Then
'MsgBox ("login succesfully")
'frmLogin.Hide
'MDIForm1.Show
'Unload Me

'Else
'MsgBox "invalid user", vbOKOnly + vbCritical, "error"
'txtID.SetFocus
'End If
'Else
'MsgBox "enter id and password"
'End If
'end sub


Private Sub login()


Module1.getconnected
Dim rs As New ADODB.Recordset
rs.Open "select * from login where  username='" & txtUserName.Text & "'", con, adOpenStatic, adLockReadOnly
If rs.RecordCount < 1 Then
MsgBox "username is invalid", vbOKOnly + vbCritical, "login"
 txtUserName.SetFocus
 Exit Sub
 Else
 If txtPassword.Text = rs!Password Then
 Unload Me
 MsgBox "login succesfully"
 Load MDIForm1
frmAnimated.Show
 Exit Sub
 Else
 MsgBox "password is invalid", vbOKOnly + vbCritical, "login"
 txtPassword.SetFocus
 Exit Sub
 End If
 End If
 Set rs = Nothing
 End Sub





Private Sub cmdCancel_Click()
 End
End Sub

Private Sub cmdOK_Click()
If txtUserName.Text = " " Then
MsgBox "userusername is empty"
txtid.SetFocus
Exit Sub
ElseIf txtPassword.Text = "" Then
MsgBox "password is empty"
txtPassword.SetFocus
Exit Sub
Else
Call login
End If


End Sub



Private Sub Command1_Click()
Shell "explorer http://www.honda2wheelersindia.com/"
End Sub

Private Sub Command2_Click()
Shell "explorer http://m.facebook.com/MY-WINGS-HONDA-96145746391186"
End Sub

Private Sub Form_Load()
Module1.getconnected
Dim rs As New ADODB.Recordset

End Sub


Private Sub Label1_Click()
Me.Hide
frmfpass.Show
End Sub
