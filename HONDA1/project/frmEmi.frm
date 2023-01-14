VERSION 5.00
Begin VB.Form frmEmi 
   Caption         =   "EMI  Options "
   ClientHeight    =   8460
   ClientLeft      =   2490
   ClientTop       =   2265
   ClientWidth     =   13095
   Icon            =   "frmEmi.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8460
   ScaleWidth      =   13095
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      Caption         =   "EMI Details"
      ForeColor       =   &H8000000E&
      Height          =   975
      Left            =   480
      TabIndex        =   20
      Top             =   1200
      Width           =   3495
      Begin VB.ComboBox cmbet 
         Height          =   315
         ItemData        =   "frmEmi.frx":30E8C
         Left            =   1320
         List            =   "frmEmi.frx":30E99
         TabIndex        =   21
         Text            =   "Select EMI Type"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "EMI Types "
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000080&
      Caption         =   "EMI Calulator"
      ForeColor       =   &H8000000E&
      Height          =   4335
      Left            =   360
      TabIndex        =   6
      Top             =   2280
      Width           =   6975
      Begin VB.TextBox txtloan 
         Height          =   375
         Left            =   1440
         TabIndex        =   11
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox txtint 
         Height          =   375
         Left            =   1440
         TabIndex        =   10
         Top             =   960
         Width           =   2295
      End
      Begin VB.TextBox txtyr 
         Height          =   375
         Left            =   1200
         TabIndex        =   9
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Submit"
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   615
         Left            =   2880
         TabIndex        =   7
         Top             =   3600
         Width           =   2535
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount of Intrest"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount of Loan "
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Intrest Rate"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Monthly Payments"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label Lblpayment 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   1560
         TabIndex        =   15
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Paid Amount"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label lblpa 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   1680
         TabIndex        =   13
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1680
         TabIndex        =   12
         Top             =   1560
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00000080&
      Caption         =   "Customer Details"
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   5400
      TabIndex        =   3
      Top             =   1200
      Width           =   4335
      Begin VB.ComboBox cmb123 
         Height          =   315
         ItemData        =   "frmEmi.frx":30EB9
         Left            =   1680
         List            =   "frmEmi.frx":30EBB
         TabIndex        =   4
         Text            =   "Select Customer"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
         ForeColor       =   &H8000000E&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00000080&
      Caption         =   "Loan on Bike "
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   7680
      TabIndex        =   0
      Top             =   2520
      Width           =   3255
      Begin VB.ComboBox cmbbike 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Bike Name"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.Image Image1 
      Height          =   32400
      Left            =   -12000
      Picture         =   "frmEmi.frx":30EBD
      Top             =   -7200
      Width           =   57600
   End
End
Attribute VB_Name = "frmEmi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim N As Integer
Dim amt, payment, rate As Double



Private Sub cmbbike_Click()
Module1.retdata ("select bprice from bike where bname ='" & cmbbike.Text & "'")
If Not rs1.EOF Or Not rs1.BOF Then
txtloan.Text = rs1.Fields(0)

End If
End Sub

Private Sub Command1_Click()
amt = Val(txtloan.Text)
rate = Val(txtloan.Text) * (Val(txtint.Text) / 100)

Label9.Caption = rate
N = Val(txtyr.Text) * 12
payment = (rate / N)
Lblpayment.Caption = Format(payment, "###0")
lblpa.Caption = rate + Val(txtloan.Text)
End Sub

Private Sub Command2_Click()

If cmbet.Text <> "" And cmb123.Text <> "" And cmbbike.Text <> "" And txtloan.Text <> "" And txtint.Text <> "" And txtia.Text <> "" And txtyr.Text <> "" And Lblpayment.Caption <> "" And lblpa.Caption <> "" Then
Module1.inupdel ("insert into emi  values ('" & cmbet.Text & "','" & cmb123.Text & "','" & cmbbike.Text & "','" & txtloan.Text & "','" & txtint.Text & "','" & txtia.Text & "','" & txtyr.Text & "' ,'" & Lblpayment.Caption & "','" & lblpa.Caption & "')")
MsgBox ("Data Saved")
Else
MsgBox ("Insert all Details "), vbOKOnly + vbCritical
End If
End Sub

Private Sub Form_Load()
Module1.retdata ("select * from customer")
While Not rs1.EOF
cmb123.AddItem (rs1.Fields(1).Value)
rs1.MoveNext
Wend
Module1.retdata ("select * from bike")
While Not rs1.EOF
cmbbike.AddItem (rs1.Fields(1).Value)
rs1.MoveNext
Wend

End Sub

