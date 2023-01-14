VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmbike 
   Caption         =   "Form2"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10245
   Icon            =   "frmbike.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6510
   ScaleWidth      =   10245
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Bike Details"
      ForeColor       =   &H00FFFFFF&
      Height          =   6855
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.PictureBox pic1 
         BackColor       =   &H00000000&
         Height          =   3135
         Left            =   3600
         ScaleHeight     =   3075
         ScaleWidth      =   2835
         TabIndex        =   12
         Top             =   360
         Width           =   2895
      End
      Begin VB.TextBox txtbid 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtbname 
         Height          =   375
         Left            =   1320
         TabIndex        =   10
         Top             =   960
         Width           =   2055
      End
      Begin VB.TextBox txtbdealer 
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Text            =   "MY WINGS HONDA"
         Top             =   1800
         Width           =   2055
      End
      Begin VB.TextBox txtbprice 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Top             =   2400
         Width           =   1455
      End
      Begin VB.ComboBox cmb1 
         Height          =   315
         ItemData        =   "frmbike.frx":30E8C
         Left            =   1440
         List            =   "frmbike.frx":30EB1
         TabIndex        =   7
         Text            =   "None"
         Top             =   2880
         Width           =   1575
      End
      Begin VB.ComboBox cmb2 
         Height          =   315
         ItemData        =   "frmbike.frx":30F19
         Left            =   1320
         List            =   "frmbike.frx":30F23
         TabIndex        =   6
         Text            =   "None"
         Top             =   3480
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add New"
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   5040
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   735
         Left            =   1920
         TabIndex        =   4
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Back"
         Height          =   735
         Left            =   3600
         TabIndex        =   3
         Top             =   5040
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cancel "
         Height          =   735
         Left            =   5160
         TabIndex        =   2
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Upload"
         Height          =   615
         Left            =   4080
         TabIndex        =   1
         Top             =   4200
         Width           =   1815
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   4680
         Top             =   4320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lblo 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   1560
         TabIndex        =   19
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bike id"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bike Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Dealer Name"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Color"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Catagory "
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   3480
         Width           =   975
      End
   End
   Begin VB.Image Image2 
      Height          =   12120
      Left            =   0
      Picture         =   "frmbike.frx":30F3C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20280
   End
   Begin VB.Image Image1 
      Height          =   32400
      Left            =   -10440
      Top             =   -12720
      Width           =   57600
   End
End
Attribute VB_Name = "frmbike"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ImageSorce As String


Private Sub Command1_Click()

txtbname.Text = ""
txtbdealer.Text = "MY WINGS HONDA"
txtbprice.Text = ""
cmb1.Text = "None"
cmb2.Text = "None"
pic1.Picture = New StdPicture

 Module1.numval ("select max(bikeid)+1 from bike")
 txtbid.Text = nm
txtbdealer.Enabled = False

End Sub

Private Sub Command2_Click()
If txtbid.Text <> "" And txtbname.Text <> "" And txtbdealer.Text <> "" And txtbprice.Text <> "" And cmb1.Text <> "" And cmb2.Text <> "" And pic1.Picture Then
Module1.inupdel ("insert into bike  values ('" & txtbid.Text & "','" & txtbname.Text & "','" & txtbdealer.Text & "','" & txtbprice.Text & "','" & cmb1.Text & "','" & cmb2.Text & "','" & txtbid.Text + ".jpg" & "' )")
Dim destPath As String
destPath = DataPath & txtbid.Text & ".jpg"
FileCopy ImageSorce, destPath
MsgBox ("Data Saved")
Module1.retdata ("select * from bike")
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

Private Sub Command5_Click()
CommonDialog1.ShowOpen
CommonDialog1.Filter = "jpeg|*.jpg"
str = CommonDialog1.FileName
ImageSorce = CommonDialog1.FileName
pic1.Picture = LoadPicture(str)
End Sub

Private Sub Form_Load()
Module1.getconnected
Module1.inupdel ("select * from bike")
ImageSorce = ""
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub txtbname_KeyPress(KeyAscii As Integer)
If KeyAscii > 65 And KeyAscii <= 90 Or KeyAscii >= 97 And KeyAscii <= 122 Or KeyAscii = 32 Or KeyAscii = 8 Then
Else
 lblo.Caption = "enter charecter"
KeyAscii = 0
txtbname.SetFocus
lblo.ForeColor = &HFF&
End If
End Sub
