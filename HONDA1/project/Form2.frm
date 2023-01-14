VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEmp 
   Caption         =   " "
   ClientHeight    =   7230
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9495
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7230
   ScaleWidth      =   9495
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12480
      TabIndex        =   14
      Top             =   4320
      Width           =   255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000012&
      Caption         =   "Employee Details"
      ForeColor       =   &H00FFFFFF&
      Height          =   6495
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   6975
      Begin VB.PictureBox pic1 
         BackColor       =   &H00FFFFFF&
         Height          =   3135
         Left            =   3960
         ScaleHeight     =   3075
         ScaleWidth      =   2835
         TabIndex        =   15
         Top             =   600
         Width           =   2895
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Upload"
         Height          =   735
         Left            =   5040
         TabIndex        =   9
         Top             =   4200
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Back"
         Height          =   495
         Left            =   3240
         TabIndex        =   8
         Top             =   5280
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   4920
         TabIndex        =   7
         Top             =   5280
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save "
         Height          =   495
         Left            =   1680
         TabIndex        =   6
         Top             =   5280
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "ADD NEW"
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   5280
         Width           =   1215
      End
      Begin VB.TextBox txteNo 
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox txteadd 
         Height          =   975
         Left            =   1680
         TabIndex        =   3
         Top             =   2400
         Width           =   2055
      End
      Begin VB.TextBox txtename 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox txteid 
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5400
         Top             =   4320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Labo 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   1680
         TabIndex        =   17
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Label lblo 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   1680
         TabIndex        =   16
         Top             =   1680
         Width           =   1935
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Contact No"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   3360
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Address"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name "
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID "
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   12000
      Left            =   0
      Picture         =   "Form2.frx":30E8C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20280
   End
End
Attribute VB_Name = "frmEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ImageSorce As String
Private Sub Command1_Click()

txtename.Text = ""
txteadd.Text = ""
txteNo.Text = ""
Module1.numval ("select max(empid)+1 from Employee")

txteid.Text = nm

End Sub

Private Sub Command2_Click()
If txteid.Text <> "" And txtename.Text <> "" And txteadd.Text <> "" And txteNo.Text <> "" And pic1.Picture Then
Module1.inupdel ("insert into employee values('" & txteid.Text & "','" & txtename.Text & "','" & txteadd.Text & "','" & txteNo.Text & "','" & txteid.Text + ".jpg" & "' )")
Dim destPath As String
destPath = DPath & txteid.Text & ".jpg"
FileCopy ImageSorce, destPath
MsgBox ("Data Saved")
Module1.retdata ("select * from employee")
Else
MsgBox ("Isert all value"), vbOKOnly + vbCritical
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
Module1.inupdel ("select * from employee")

ImageSorce = ""

End Sub



Private Sub txtename_KeyPress(KeyAscii As Integer)
If KeyAscii > 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32 Or KeyAscii = 8 Then
Else
 lblo.Caption = "enter charecter"
KeyAscii = 0
txtename.SetFocus
lblo.ForeColor = &HFF&
End If
End Sub

Private Sub txteNo_LostFocus()
If Len(txteNo.Text) = 10 Then
Labo.Caption = "Valid No"
Labo.ForeColor = &H8000&
Else
Labo.Caption = "Invalid No"
Labo.ForeColor = &HFF&
End If
End Sub
