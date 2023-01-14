VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmubk 
   Caption         =   "Form4"
   ClientHeight    =   7080
   ClientLeft      =   3270
   ClientTop       =   2070
   ClientWidth     =   9510
   Icon            =   "frmubk.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   7080
   ScaleWidth      =   9510
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command5 
      Caption         =   "Upload"
      Height          =   735
      Left            =   4080
      TabIndex        =   11
      Top             =   5040
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "UPDATE"
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   5520
      Width           =   2295
   End
   Begin VB.TextBox txtbname 
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtsrh 
      Height          =   615
      Left            =   1800
      TabIndex        =   8
      Top             =   0
      Width           =   3855
   End
   Begin VB.ComboBox cmb2 
      Height          =   315
      ItemData        =   "frmubk.frx":30E8C
      Left            =   1440
      List            =   "frmubk.frx":30E96
      TabIndex        =   7
      Text            =   "None"
      Top             =   4200
      Width           =   1455
   End
   Begin VB.ComboBox cmb1 
      Height          =   315
      ItemData        =   "frmubk.frx":30EAF
      Left            =   1440
      List            =   "frmubk.frx":30ED4
      TabIndex        =   6
      Text            =   "None"
      Top             =   3480
      Width           =   1575
   End
   Begin VB.TextBox txtbprice 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtbdealer 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Text            =   "MY WINGS HONDA"
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox txtbid 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.PictureBox pic12 
      BackColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   4080
      ScaleHeight     =   3075
      ScaleWidth      =   4275
      TabIndex        =   2
      Top             =   1080
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
      Height          =   735
      Left            =   4080
      TabIndex        =   1
      Top             =   6000
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update Photo"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   6240
      Width           =   2295
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Catagory "
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4320
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Color"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dealer Name"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   2280
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bike Name"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bike id"
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Search from name"
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   18000
      Left            =   0
      Picture         =   "frmubk.frx":30F3C
      Top             =   0
      Width           =   28800
   End
End
Attribute VB_Name = "frmubk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ImageSorce As String
Private Sub Combo1_Click()
txtbid.Enabled = False
txtbname.Enabled = False
txtbdealer.Enabled = False
txtbprice.Enabled = False
cmb1.Enabled = False
cmb2.Enabled = False
pic12.Enabled = False
End Sub

Private Sub Command1_Click()
  Me.Hide
  MDIForm1.Show
End Sub

Private Sub Command2_Click()
Module1.inupdel ("update bike set bike.bname = '" & txtbname.Text & "', bike.bprice = '" & txtbprice.Text & "', bike.Dealername = '" & txtbdealer.Text & "' ,bike.BikeColor = '" & cmb1.Text & "',bike.BikeCatagory = '" & cmb2.Text & "' where bike.bikeid = '" & txtbid.Text & "'")
MsgBox "Details Updated ", vbInformation
End Sub

Private Sub Command3_Click()
Module1.inupdel ("update bike set bike.photo = '" & txtbid.Text + ".jpg" & "' where bike.bikeid = '" & txtbid.Text & "'")
Dim destPath As String
destPath = DataPath & txtbid.Text & ".jpg"
FileCopy ImageSorce, destPath
MsgBox "Details Updated ", vbInformation
End Sub

Private Sub Command5_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "jpeg|*.jpg"
str = CommonDialog1.FileName
ImageSorce = CommonDialog1.FileName
pic12.Picture = LoadPicture(str)
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
Private Sub loaddata()
On Error Resume Next
txtbid.Text = rs1.Fields(0).Value
txtbname.Text = rs1.Fields(1).Value
txtbdealer.Text = rs1.Fields(3).Value
txtbprice.Text = rs1.Fields(2).Value
cmb1.Text = rs1.Fields(4).Value
cmb2.Text = rs1.Fields(5).Value
pic12.Picture = LoadPicture(DataPath & rs1.Fields(6).Value)

End Sub

