VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmuc 
   Caption         =   "Form5"
   ClientHeight    =   8295
   ClientLeft      =   2055
   ClientTop       =   675
   ClientWidth     =   10365
   Icon            =   "frmuc.frx":0000
   LinkTopic       =   "Form5"
   ScaleHeight     =   8295
   ScaleWidth      =   10365
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Update"
      ForeColor       =   &H8000000B&
      Height          =   9015
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7815
      Begin VB.CommandButton Command1 
         Caption         =   "UPDATE"
         Height          =   615
         Left            =   960
         TabIndex        =   11
         Top             =   7200
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Back"
         Height          =   615
         Left            =   5160
         TabIndex        =   10
         Top             =   7200
         Width           =   1455
      End
      Begin VB.TextBox txtcmail 
         Height          =   495
         Left            =   1560
         TabIndex        =   9
         Top             =   5760
         Width           =   2415
      End
      Begin VB.TextBox txtcNo 
         Height          =   375
         Left            =   1560
         TabIndex        =   8
         Top             =   5160
         Width           =   1695
      End
      Begin VB.TextBox txtcAdd 
         Height          =   1455
         Left            =   1560
         TabIndex        =   7
         Top             =   3240
         Width           =   2655
      End
      Begin VB.TextBox txtcName 
         Height          =   375
         Left            =   1560
         TabIndex        =   6
         Top             =   2520
         Width           =   2055
      End
      Begin VB.TextBox txtcid 
         Height          =   405
         Left            =   1560
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox txtsch 
         Height          =   495
         Left            =   1560
         TabIndex        =   4
         Top             =   480
         Width           =   3015
      End
      Begin VB.PictureBox pic1 
         BackColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   4560
         ScaleHeight     =   3075
         ScaleWidth      =   2835
         TabIndex        =   3
         Top             =   1680
         Width           =   2895
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Upload"
         Height          =   735
         Left            =   5280
         TabIndex        =   2
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Update Photo"
         Height          =   615
         Left            =   960
         TabIndex        =   1
         Top             =   8040
         Width           =   2415
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5760
         Top             =   3840
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Email Adress"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   5760
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Search by name"
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   360
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Custmer Phone No"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   5160
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Custmer Adsress"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Custmer Name"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Custmer Id"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   18000
      Left            =   120
      Picture         =   "frmuc.frx":30E8C
      Top             =   -3240
      Width           =   28800
   End
End
Attribute VB_Name = "frmuc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ImageSorce As String


Private Sub Command1_Click()
Module1.inupdel ("update customer set customer.custname = '" & txtcName.Text & "',customer.customeraddress = '" & txtcAdd.Text & "',customer.customerphone = '" & txtcNo.Text & "',customer.EmailId = '" & txtcmail.Text & "'where customer.customerid = '" & txtcid.Text & "'")
MsgBox "Details Updated ", vbInformation
End Sub

Private Sub Command2_Click()
Me.Hide
MDIForm1.Show
End Sub

Private Sub Command3_Click()

Module1.inupdel ("update customer set customer.photo='" & txtcid.Text + ".jpg" & "' where customer.customerid = '" & txtcid.Text & "'")
Dim destPath As String
destPath = DPath & txtcid.Text & ".jpg"
FileCopy ImageSorce, destPath
MsgBox "Details Updated ", vbInformation

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



