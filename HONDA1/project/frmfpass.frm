VERSION 5.00
Begin VB.Form frmfpass 
   Caption         =   "Form2"
   ClientHeight    =   8790
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13680
   Icon            =   "frmfpass.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8790
   ScaleWidth      =   13680
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Forgot Password"
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7695
      Begin VB.CommandButton changepassbtn 
         Caption         =   "Change Password"
         Height          =   1095
         Left            =   2160
         TabIndex        =   7
         Top             =   6360
         Width           =   2415
      End
      Begin VB.TextBox txtconfirm 
         Height          =   1095
         Left            =   1920
         TabIndex        =   6
         Top             =   5040
         Width           =   3855
      End
      Begin VB.TextBox txtnew 
         Height          =   615
         Left            =   1920
         TabIndex        =   5
         Top             =   4080
         Width           =   3615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Verify"
         Height          =   615
         Left            =   5760
         TabIndex        =   4
         Top             =   1680
         Width           =   1455
      End
      Begin VB.TextBox txtfdob 
         Height          =   495
         Left            =   1920
         TabIndex        =   3
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox txtfuser 
         Height          =   495
         Left            =   1920
         TabIndex        =   2
         Top             =   480
         Width           =   3255
      End
      Begin VB.CommandButton cmdcheak 
         Caption         =   "Cheak"
         Height          =   615
         Left            =   5760
         TabIndex        =   1
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblcp 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password "
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Label lblenp 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter New Password"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   4200
         Width           =   1575
      End
      Begin VB.Label lbl3 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   1560
         TabIndex        =   11
         Top             =   2640
         Width           =   3855
      End
      Begin VB.Label lbl2 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   1200
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Birth"
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Image Image1 
      Height          =   16200
      Left            =   0
      Picture         =   "frmfpass.frx":30E8C
      Top             =   0
      Width           =   28800
   End
End
Attribute VB_Name = "frmfpass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub changepassbtn_Click()
If txtnew.Text = txtconfirm.Text Then
Module1.inupdel ("update login set login.password = '" & txtconfirm.Text & "'  where login.username = '" & txtfuser.Text & "'")
MsgBox "Password Changed Successfully", vbInformation, "Password Change: Success"
Me.Hide
frmLogin.Show
Else
MsgBox "Password Does not matched,Please Enter Correct Details", vbExclamation, "Change Password:Failed"
txtnew.Text = ""
txtconfirm.Text = ""
End If
End Sub

Private Sub cmdcheak_Click()
rs.Open "select * from login where  username='" & txtfuser.Text & "'", con, adOpenStatic, adLockReadOnly
If rs.RecordCount < 1 Then
lbl2.Caption = "User id not found .....Sorry can't Reset password!!"
txtfuser.SetFocus
lbl2.ForeColor = &HFF&
Else
lbl2.Caption = "User ID Found in the database"
lbl2.ForeColor = &H8000&

End If
 Set rs = Nothing
End Sub

Private Sub Command1_Click()
rs.Open "select * from login where  dob='" & txtfdob.Text & "'", con, adOpenStatic, adLockReadOnly
If rs.RecordCount < 1 Then
lbl3.Caption = "Account not verified ,Can't reset the password"
lbl2.Caption = "Sorry .. Date of Birth Not Matched !! "
lbl2.ForeColor = &HFF&
lbl3.ForeColor = &HFF&

Else

lbl2.ForeColor = &H8000&
lbl3.ForeColor = &H8000&
lbl2.Caption = "Congratulations !!"
lbl3.Caption = "Account is verified Now,Set your new Password"
txtnew.Visible = True
txtconfirm.Visible = True
changepassbtn.Visible = True
lblenp.Visible = True
lblcp.Visible = True
End If
 Set rs = Nothing
End Sub

 Private Sub Form_Load()
txtnew.Visible = False
txtconfirm.Visible = False
changepassbtn.Visible = False
lblenp.Visible = False
lblcp.Visible = False

End Sub
