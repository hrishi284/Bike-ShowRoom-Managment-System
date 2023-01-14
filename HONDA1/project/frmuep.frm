VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmuep 
   Caption         =   "Form3"
   ClientHeight    =   7335
   ClientLeft      =   4965
   ClientTop       =   840
   ClientWidth     =   9240
   Icon            =   "frmuep.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   7335
   ScaleWidth      =   9240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Update"
      ForeColor       =   &H8000000B&
      Height          =   8415
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   8295
      Begin VB.CommandButton Command5 
         Caption         =   "Upload"
         Height          =   735
         Left            =   5760
         TabIndex        =   10
         Top             =   5520
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "UPDATE"
         Height          =   615
         Left            =   1320
         TabIndex        =   9
         Top             =   6480
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Back"
         Height          =   735
         Left            =   3960
         TabIndex        =   8
         Top             =   6480
         Width           =   1455
      End
      Begin VB.TextBox txtid 
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtname 
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   2880
         Width           =   2055
      End
      Begin VB.TextBox txtadd 
         Height          =   975
         Left            =   2400
         TabIndex        =   5
         Top             =   3600
         Width           =   2415
      End
      Begin VB.TextBox txtno 
         Height          =   375
         Left            =   2400
         TabIndex        =   4
         Top             =   5400
         Width           =   2055
      End
      Begin VB.TextBox txtsch 
         Height          =   495
         Left            =   2400
         TabIndex        =   3
         Top             =   1080
         Width           =   3975
      End
      Begin VB.PictureBox pic1 
         BackColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   5040
         ScaleHeight     =   3075
         ScaleWidth      =   2835
         TabIndex        =   2
         Top             =   1920
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Update Photo"
         Height          =   615
         Left            =   1320
         TabIndex        =   1
         Top             =   7320
         Width           =   2295
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5760
         Top             =   5640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000012&
         Caption         =   "Employee ID "
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   720
         TabIndex        =   15
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000012&
         Caption         =   "Employee Name "
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000012&
         Caption         =   "Employee Address"
         ForeColor       =   &H8000000B&
         Height          =   255
         Left            =   480
         TabIndex        =   13
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000012&
         Caption         =   "Employee Contact No"
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000012&
         Caption         =   "Search by name"
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   720
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   18000
      Left            =   -240
      Picture         =   "frmuep.frx":30E8C
      Top             =   -2280
      Width           =   28800
   End
End
Attribute VB_Name = "frmuep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ImageSorce As String
Private Sub Command1_Click()
Me.Hide
  MDIForm1.Show

End Sub

Private Sub Command3_Click()
Module1.inupdel ("update employee set employee.photo='" & txtid.Text + ".jpg" & "' where employee.empid = '" & txtid.Text & "'")
Dim destPath As String
destPath = DataP & txtid.Text & ".jpg"
FileCopy ImageSorce, destPath
MsgBox "Details Updated ", vbInformation
End Sub





Private Sub Command2_Click()
Module1.inupdel ("update employee set employee.empname = '" & txtname.Text & "',employee.empAdress = '" & txtadd.Text & "',employee.empNo = '" & txtno.Text & "'  where employee.empid = '" & txtid.Text & "'")
MsgBox "Details Updated ", vbInformation
End Sub

Private Sub Command5_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "jpeg|*.jpg"
str = CommonDialog1.FileName
ImageSorce = CommonDialog1.FileName
pic1.Picture = LoadPicture(str)
End Sub

Private Sub txtsch_KeyPress(KeyAscii As Integer)
Module1.retdata ("select * from employee where empname like '%" & txtsch.Text & "%' ")

If Not rs1.BOF Or Not rs1.EOF Then
Call loaddata
Else
txtid.Text = ""
txtname.Text = ""
txtadd.Text = ""
txtno.Text = ""
pic1.Picture = New StdPicture
MsgBox ("No Data Found")
End If
End Sub

Private Sub loaddata()
On Error Resume Next
txtid.Text = rs1.Fields(0).Value
txtname.Text = rs1.Fields(1).Value
txtadd.Text = rs1.Fields(2).Value
txtno.Text = rs1.Fields(3).Value
ImageSorce = CommonDialog1.FileName
pic1.Picture = LoadPicture(DataP & rs1.Fields(4).Value)
End Sub



