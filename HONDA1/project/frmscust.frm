VERSION 5.00
Begin VB.Form frmscust 
   ClientHeight    =   6270
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8910
   Icon            =   "frmscust.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6270
   ScaleWidth      =   8910
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Custmer Seach"
      Height          =   7095
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   7575
      Begin VB.PictureBox pic1 
         BackColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   4440
         ScaleHeight     =   3075
         ScaleWidth      =   2835
         TabIndex        =   15
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   615
         Left            =   2280
         TabIndex        =   8
         Top             =   6240
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Back"
         Height          =   615
         Left            =   480
         TabIndex        =   7
         Top             =   6240
         Width           =   1455
      End
      Begin VB.TextBox txtsch 
         Height          =   495
         Left            =   1680
         TabIndex        =   6
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox txtcid 
         Height          =   405
         Left            =   1680
         TabIndex        =   5
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtcName 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   2160
         Width           =   2055
      End
      Begin VB.TextBox txtcAdd 
         Height          =   1455
         Left            =   1680
         TabIndex        =   3
         Top             =   2760
         Width           =   2655
      End
      Begin VB.TextBox txtcNo 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   4440
         Width           =   1695
      End
      Begin VB.TextBox txtcmail 
         Height          =   495
         Left            =   1680
         TabIndex        =   1
         Top             =   5160
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Custmer Id"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Custmer Name"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Custmer Adsress"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   2880
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Custmer Phone No"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   4440
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Email Adress"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   5280
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Search by name"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   10965
      Left            =   -360
      Picture         =   "frmscust.frx":30E8C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20640
   End
End
Attribute VB_Name = "frmscust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
End
End Sub



Private Sub Command3_Click()

Dim rn As String
If MsgBox("Do You Want To Delete?", vbQuestion + vbYesNo) = vbYes Then
  rn = InputBox("Enter Name to Delete")
  Module1.retdata ("select * from customer where custname = '" & rn & "'")
 
  If Not rs1.EOF And Not rs1.BOF Then
   Module1.inupdel ("delete * from customer where custname = '" & rn & "'")
   MsgBox ("Record was Deleted ")
   Else
   MsgBox ("Record not exist"), vbInformation
   End If
Else
MsgBox ("operation Cancelled")
End If

End Sub

Private Sub Command4_Click()
Me.Hide
MDIForm1.Show
End Sub

Private Sub txtsch_KeyPress(KeyAscii As Integer)
Module1.retdata ("select * from customer where custname like '%" & txtsch.Text & "%' ")

If Not rs1.BOF Or Not rs1.EOF Then
Call loaddata
Else
txtcid.Text = ""
txtcName.Text = ""
txtcAdd.Text = ""
txtcNo.Text = ""
txtcmail.Text = ""
pic1.Picture = New StdPicture
MsgBox ("No Data Found")
End If
End Sub

Private Sub loaddata()
On Error Resume Next
txtcid.Text = rs1.Fields(0).Value
txtcName.Text = rs1.Fields(1).Value
txtcAdd.Text = rs1.Fields(2).Value
txtcNo.Text = rs1.Fields(3).Value
txtcmail.Text = rs1.Fields(4).Value
pic1.Picture = LoadPicture(DPath & rs1.Fields(6).Value)
End Sub

