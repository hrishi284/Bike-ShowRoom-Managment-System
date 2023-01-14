VERSION 5.00
Begin VB.Form frmenq 
   Caption         =   "Form2"
   ClientHeight    =   4290
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6930
   Icon            =   "frmenq.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   4290
   ScaleWidth      =   6930
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Enquiry Details"
      ForeColor       =   &H8000000B&
      Height          =   7935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
      Begin VB.ComboBox cmb1 
         Height          =   315
         Left            =   2400
         TabIndex        =   11
         Text            =   "Select Bike"
         Top             =   5520
         Width           =   1575
      End
      Begin VB.TextBox txtNo 
         Height          =   375
         Left            =   2400
         TabIndex        =   10
         Top             =   840
         Width           =   2295
      End
      Begin VB.TextBox txtdt 
         Height          =   375
         Left            =   2400
         TabIndex        =   9
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtcname 
         Height          =   285
         Left            =   2400
         TabIndex        =   8
         Top             =   2040
         Width           =   2775
      End
      Begin VB.TextBox txtph 
         Height          =   375
         Left            =   2400
         TabIndex        =   7
         Top             =   2640
         Width           =   2415
      End
      Begin VB.TextBox txtadd 
         Height          =   1215
         Left            =   2400
         TabIndex        =   6
         Top             =   3240
         Width           =   3015
      End
      Begin VB.TextBox txteid 
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   4920
         Width           =   2535
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add New"
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   6360
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   615
         Left            =   1920
         TabIndex        =   3
         Top             =   6360
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Back"
         Height          =   615
         Left            =   3600
         TabIndex        =   2
         Top             =   6360
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cancel"
         Height          =   615
         Left            =   5280
         TabIndex        =   1
         Top             =   6360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Enquiry No"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Enquiry Date"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   360
         TabIndex        =   17
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Custome rPhone"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Address"
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Email-ID"
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   240
         TabIndex        =   13
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Bike Intersted "
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   5520
         Width           =   1335
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Enquiry No"
      Height          =   375
      Left            =   480
      TabIndex        =   19
      Top             =   1560
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   16200
      Left            =   -2520
      Picture         =   "frmenq.frx":30E8C
      Top             =   -240
      Width           =   28800
   End
End
Attribute VB_Name = "frmenq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
txtno.Text = ""
txtdt.Text = ""
txtcName.Text = ""
txtph.Text = ""
txtadd.Text = ""
txteid.Text = ""
cmb1.Text = "Select bike"
End Sub



Private Sub Command2_Click()
If txtno.Text <> "" And txtdt.Text <> "" And txtcName.Text <> "" And txtph.Text <> "" And txtadd.Text <> "" And txteid.Text <> "" And cmb1.Text <> "" Then
Module1.inupdel ("insert into enquiry values ('" & txtno.Text & "','" & txtdt.Text & "','" & txtcName.Text & "','" & txtph.Text & "','" & txtadd.Text & "','" & txteid.Text & "','" & cmb1.Text & "' )")
MsgBox ("Data Saved")
Module1.retdata ("select * from enquiry")
Else
MsgBox ("Insert all Details "), vbOKOnly + vbCritical
End If
End Sub

Private Sub Command3_Click()
Me.Hide
MDIForm1.Show
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
Module1.retdata ("select * from bike")
While Not rs1.EOF
cmb1.AddItem (rs1.Fields(1).Value)
rs1.MoveNext
Wend
Module1.getconnected
Module1.inupdel ("select * from enquiry ")

End Sub
