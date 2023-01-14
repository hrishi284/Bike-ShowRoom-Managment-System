VERSION 5.00
Begin VB.Form frmsEmp 
   Caption         =   "Form2"
   ClientHeight    =   8310
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9420
   Icon            =   "frmsEmp.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8310
   ScaleWidth      =   9420
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Employee Details"
      ForeColor       =   &H8000000E&
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.TextBox txtid 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   1320
         Width           =   2055
      End
      Begin VB.TextBox txtname 
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox txtadd 
         Height          =   975
         Left            =   1560
         TabIndex        =   6
         Top             =   2760
         Width           =   2415
      End
      Begin VB.TextBox txtno 
         Height          =   375
         Left            =   1560
         TabIndex        =   5
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   2160
         TabIndex        =   4
         Top             =   5040
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Back"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   5040
         Width           =   1455
      End
      Begin VB.TextBox txtsch 
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
      Begin VB.PictureBox pic1 
         BackColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   4080
         ScaleHeight     =   3075
         ScaleWidth      =   2835
         TabIndex        =   1
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name "
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Address"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Contact No"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Search by name"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   16200
      Left            =   0
      Picture         =   "frmsEmp.frx":30E8C
      Top             =   0
      Width           =   28800
   End
End
Attribute VB_Name = "frmsEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Dim rn As String
If MsgBox("Do You Want To Delete?", vbQuestion + vbYesNo) = vbYes Then
  rn = InputBox("Enter Name to Delete")
  Module1.retdata ("select * from employee where empname = '" & rn & "'")
 
  If Not rs1.EOF And Not rs1.BOF Then
   Module1.inupdel ("delete * from employee where empname = '" & rn & "'")
   MsgBox ("Record was Deleted ")
   Else
   MsgBox ("Record not exist"), vbInformation
   End If
Else
MsgBox ("operation Cancelled")
End If
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Me.Hide
MDIForm1.Show
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
pic1.Picture = LoadPicture(DataP & rs1.Fields(4).Value)
End Sub

