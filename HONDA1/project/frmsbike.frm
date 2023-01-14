VERSION 5.00
Begin VB.Form frmsbike 
   Caption         =   "Form2"
   ClientHeight    =   8325
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   Icon            =   "frmsbike.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8325
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Caption         =   "Bike Search "
      ForeColor       =   &H00FFFFFF&
      Height          =   6855
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7815
      Begin VB.PictureBox pic12 
         BackColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   3360
         ScaleHeight     =   3075
         ScaleWidth      =   4275
         TabIndex        =   17
         Top             =   1920
         Width           =   4335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   615
         Left            =   2040
         TabIndex        =   9
         Top             =   5760
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Back"
         Height          =   615
         Left            =   240
         TabIndex        =   8
         Top             =   5760
         Width           =   1455
      End
      Begin VB.TextBox txtbid 
         Height          =   285
         Left            =   1320
         TabIndex        =   7
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox txtbname 
         Height          =   375
         Left            =   1320
         TabIndex        =   6
         Top             =   2760
         Width           =   2055
      End
      Begin VB.TextBox txtbdealer 
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Text            =   "MY WINGS HONDA"
         Top             =   3360
         Width           =   2055
      End
      Begin VB.TextBox txtbprice 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   3840
         Width           =   1455
      End
      Begin VB.ComboBox cmb1 
         Height          =   315
         ItemData        =   "frmsbike.frx":30E8C
         Left            =   1320
         List            =   "frmsbike.frx":30EB1
         TabIndex        =   3
         Text            =   "None"
         Top             =   4320
         Width           =   1575
      End
      Begin VB.ComboBox cmb2 
         Height          =   315
         ItemData        =   "frmsbike.frx":30F19
         Left            =   1320
         List            =   "frmsbike.frx":30F23
         TabIndex        =   2
         Text            =   "None"
         Top             =   4920
         Width           =   1455
      End
      Begin VB.TextBox txtsrh 
         Height          =   615
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   3855
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Search from name"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bike id"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bike Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3840
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Catagory "
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   4920
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   12000
      Left            =   -360
      Picture         =   "frmsbike.frx":30F3C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20280
   End
End
Attribute VB_Name = "frmsbike"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False








Private Sub Command1_Click()
Me.Hide
MDIForm1.Show
End Sub

Private Sub Command2_Click()
End
End Sub



Private Sub loaddata()
On Error Resume Next
txtbid.Text = rs1.Fields(0).Value
txtbname.Text = rs1.Fields(1).Value
txtbdealer.Text = rs1.Fields(2).Value
txtbprice.Text = rs1.Fields(3).Value
cmb1.Text = rs1.Fields(4).Value
cmb2.Text = rs1.Fields(5).Value
pic12.Picture = LoadPicture(DataPath & rs1.Fields(6).Value)
End Sub






Private Sub Command3_Click()
Dim rss As String
If MsgBox("Do You Want To Delete?", vbQuestion + vbYesNo) = vbYes Then
  rss = InputBox("Enter Name to Delete")
  Module1.retdata ("select * from bike where bname= '" & rss & "'")
  If Not rs1.EOF And Not rs1.BOF Then
   Module1.inupdel ("Delete * from bike where bname='" & rss & "'")
   MsgBox ("Record was Deleted ")
   Else
   MsgBox ("Record not exist"), vbInformation
   End If
Else
MsgBox ("operation Cancelled")
End If
End Sub

Private Sub txtsrh_KeyPress(KeyAscii As Integer)
Module1.retdata ("select * from bike where bname like '%" & txtsrh.Text & "%' ")
If Not rs1.BOF Or Not rs1.EOF Then
Call loaddata
Else
txtbid.Text = ""
txtbname.Text = ""
txtbdealer.Text = "MY WINGS HONDA"
txtbprice.Text = ""
cmb1.Text = "None"
cmb2.Text = "None"
pic12.Picture = New StdPicture
MsgBox ("No Data Found")
End If
End Sub
