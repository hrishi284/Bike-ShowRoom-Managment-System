VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCust 
   Caption         =   "Form2"
   ClientHeight    =   7800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10200
   Icon            =   "frmCust.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7800
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Custmer Details"
      ForeColor       =   &H80000008&
      Height          =   7335
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   7935
      Begin VB.PictureBox pic1 
         BackColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   4800
         ScaleHeight     =   3075
         ScaleWidth      =   2835
         TabIndex        =   18
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox txtadhar 
         Height          =   405
         Left            =   2040
         TabIndex        =   17
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Back"
         Height          =   615
         Left            =   3480
         TabIndex        =   10
         Top             =   6120
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cancel"
         Height          =   615
         Left            =   5280
         TabIndex        =   9
         Top             =   6120
         Width           =   1455
      End
      Begin VB.TextBox txtcmail 
         Height          =   495
         Left            =   1800
         TabIndex        =   8
         Top             =   4680
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "SAVE"
         Height          =   615
         Left            =   1800
         TabIndex        =   7
         Top             =   6120
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ADD"
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   6120
         Width           =   1215
      End
      Begin VB.TextBox txtcNo 
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox txtcAdd 
         Height          =   1455
         Left            =   1920
         TabIndex        =   4
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox txtcName 
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox txtcid 
         Height          =   285
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Upload "
         Height          =   495
         Left            =   5640
         TabIndex        =   1
         Top             =   3960
         Width           =   1335
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   6240
         Top             =   3960
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Labo 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   4080
         TabIndex        =   21
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label lblo 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2040
         TabIndex        =   20
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1800
         TabIndex        =   19
         Top             =   5280
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   " Adhar card no"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   3360
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Custmer Phone No"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   4080
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Email Adress"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   4800
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Custmer Adsress"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         TabIndex        =   13
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Custmer Id"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Custmer Name"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   16005
      Left            =   -240
      Picture         =   "frmCust.frx":30E8C
      Top             =   0
      Width           =   23415
   End
End
Attribute VB_Name = "frmCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ImageSorce As String
Private Sub Command1_Click()
txtcid.Text = ""
txtcName.Text = ""
txtcAdd.Text = ""
txtcNo.Text = ""
 Module1.numval ("select max(customerid)+1 from customer")
txtcid.Text = nm
End Sub

Private Sub Command2_Click()
If txtcid.Text <> "" And txtcName.Text <> "" And txtcAdd.Text <> "" And txtadhar.Text <> "" And txtcNo.Text <> "" And txtcmail.Text <> "" And pic1.Picture Then
Module1.inupdel ("insert into customer values ('" & txtcid.Text & "','" & txtcName.Text & "','" & txtcAdd.Text & "','" & txtadhar.Text & "','" & txtcNo.Text & "','" & txtcmail.Text & "','" & txtcid.Text + ".jpg" & "'  )")
Dim destPath As String
destPath = DPath & txtcid.Text & ".jpg"
FileCopy ImageSorce, destPath
MsgBox ("Data Saved")
Module1.retdata ("select * from customer")
Else
MsgBox ("Insert all Details "), vbOKOnly + vbCritical
End If
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Me.Hide
MDIForm1.Show
End Sub

Private Sub Command5_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "jpeg|*.jpg"
str = CommonDialog1.FileName
ImageSorce = CommonDialog1.FileName
pic1.Picture = LoadPicture(str)
End Sub

Private Sub Form_Load()
Module1.getconnected
Module1.inupdel ("select * from customer")
ImageSorce = ""
End Sub

Private Sub txtcmail_LostFocus()
If InStr(2, Trim(txtcmail.Text), "@") Then
Label.Caption = "Vailid Email"
Label.ForeColor = &H8000&
Else
Label.Caption = "invalid email"
Label.ForeColor = &HFF&
End If
End Sub


Private Sub txtcName_KeyPress(KeyAscii As Integer)
If KeyAscii > 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32 Or KeyAscii = 8 Then
Else
 lblo.Caption = "enter charecter"
KeyAscii = 0
txtcName.SetFocus
lblo.ForeColor = &HFF&
End If

End Sub

Private Sub txtcNo_LostFocus()
If Len(txtcNo.Text) = 10 Then
Labo.Caption = "Valid No"
Labo.ForeColor = &H8000&
Else
Labo.Caption = "Invalid No"
Labo.ForeColor = &HFF&
End If
End Sub
