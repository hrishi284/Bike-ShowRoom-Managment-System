VERSION 5.00
Begin VB.Form frmbill 
   BackColor       =   &H8000000E&
   Caption         =   "Form2"
   ClientHeight    =   8940
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11535
   Icon            =   "frmbill.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   8940
   ScaleWidth      =   11535
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   "Back"
      Height          =   735
      Left            =   11520
      TabIndex        =   69
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton print1 
      Caption         =   "Print "
      Height          =   735
      Left            =   13560
      TabIndex        =   67
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add Without Accessries"
      Height          =   735
      Left            =   13440
      TabIndex        =   66
      Top             =   6000
      Width           =   1575
   End
   Begin VB.Frame lblqtylblqty 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Billing "
      Height          =   10695
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   11295
      Begin VB.TextBox Text1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1320
         TabIndex        =   70
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Lbldate 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   7560
         TabIndex        =   68
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Image Image1 
         Height          =   1260
         Left            =   360
         Picture         =   "frmbill.frx":30E8C
         Top             =   480
         Width           =   1500
      End
      Begin VB.Label lbltot 
         Caption         =   "Label11"
         Height          =   255
         Left            =   8400
         TabIndex        =   65
         Top             =   9840
         Width           =   975
      End
      Begin VB.Label lblgst 
         Caption         =   "Label11"
         Height          =   255
         Left            =   8280
         TabIndex        =   64
         Top             =   9480
         Width           =   1095
      End
      Begin VB.Label lblta 
         Caption         =   "Label11"
         Height          =   375
         Left            =   9000
         TabIndex        =   63
         Top             =   9120
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "TOTAL -"
         Height          =   255
         Left            =   7440
         TabIndex        =   62
         Top             =   9840
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "GST-"
         Height          =   255
         Left            =   7440
         TabIndex        =   61
         Top             =   9480
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "TOTAL AMOUNT-"
         Height          =   255
         Left            =   7440
         TabIndex        =   60
         Top             =   9120
         Width           =   2175
      End
      Begin VB.Label lblamt 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   5
         Left            =   8760
         TabIndex        =   59
         Top             =   7800
         Width           =   1215
      End
      Begin VB.Label lblamt 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   4
         Left            =   8760
         TabIndex        =   58
         Top             =   7200
         Width           =   1335
      End
      Begin VB.Label lblamt 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   3
         Left            =   8760
         TabIndex        =   57
         Top             =   6600
         Width           =   1335
      End
      Begin VB.Label lblamt 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   2
         Left            =   8760
         TabIndex        =   56
         Top             =   6000
         Width           =   1455
      End
      Begin VB.Label lblamt 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   8760
         TabIndex        =   55
         Top             =   5400
         Width           =   1455
      End
      Begin VB.Label lblamt 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   8880
         TabIndex        =   54
         Top             =   4920
         Width           =   1215
      End
      Begin VB.Label lblqty 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   5
         Left            =   7560
         TabIndex        =   53
         Top             =   7800
         Width           =   735
      End
      Begin VB.Label lblqty 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   4
         Left            =   7560
         TabIndex        =   52
         Top             =   7200
         Width           =   735
      End
      Begin VB.Label lblqty 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   3
         Left            =   7560
         TabIndex        =   51
         Top             =   6600
         Width           =   735
      End
      Begin VB.Label lblqty 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   2
         Left            =   7560
         TabIndex        =   50
         Top             =   6000
         Width           =   735
      End
      Begin VB.Label lblqty 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   7560
         TabIndex        =   49
         Top             =   5400
         Width           =   735
      End
      Begin VB.Label lblqty 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   7560
         TabIndex        =   48
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label lblprice 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   5
         Left            =   6240
         TabIndex        =   47
         Top             =   7800
         Width           =   735
      End
      Begin VB.Label lblprice 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   4
         Left            =   6240
         TabIndex        =   46
         Top             =   7200
         Width           =   735
      End
      Begin VB.Label lblprice 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   3
         Left            =   6240
         TabIndex        =   45
         Top             =   6600
         Width           =   735
      End
      Begin VB.Label lblprice 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   2
         Left            =   6240
         TabIndex        =   44
         Top             =   6000
         Width           =   735
      End
      Begin VB.Label lblprice 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   6240
         TabIndex        =   43
         Top             =   5400
         Width           =   735
      End
      Begin VB.Label lblprice 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   6240
         TabIndex        =   42
         Top             =   4920
         Width           =   735
      End
      Begin VB.Label lblname 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   5
         Left            =   1560
         TabIndex        =   41
         Top             =   7680
         Width           =   3855
      End
      Begin VB.Label lblname 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   4
         Left            =   1560
         TabIndex        =   40
         Top             =   7080
         Width           =   3855
      End
      Begin VB.Label lblname 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   3
         Left            =   1560
         TabIndex        =   39
         Top             =   6480
         Width           =   3855
      End
      Begin VB.Label lblname 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   2
         Left            =   1440
         TabIndex        =   38
         Top             =   6120
         Width           =   3855
      End
      Begin VB.Label lblname 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   37
         Top             =   5520
         Width           =   3855
      End
      Begin VB.Label lblname 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   0
         Left            =   1440
         TabIndex        =   36
         Top             =   4920
         Width           =   3855
      End
      Begin VB.Label lblsr 
         BackStyle       =   0  'Transparent
         Height          =   375
         Index           =   5
         Left            =   720
         TabIndex        =   35
         Top             =   7680
         Width           =   495
      End
      Begin VB.Label lblsr 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   4
         Left            =   720
         TabIndex        =   34
         Top             =   7080
         Width           =   495
      End
      Begin VB.Label lblsr 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   3
         Left            =   720
         TabIndex        =   33
         Top             =   6480
         Width           =   495
      End
      Begin VB.Label lblsr 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   2
         Left            =   720
         TabIndex        =   32
         Top             =   6000
         Width           =   495
      End
      Begin VB.Label lblsr 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   720
         TabIndex        =   31
         Top             =   5520
         Width           =   495
      End
      Begin VB.Label lblsr 
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   30
         Top             =   4920
         Width           =   495
      End
      Begin VB.Line Line3 
         X1              =   7320
         X2              =   8520
         Y1              =   9000
         Y2              =   9000
      End
      Begin VB.Shape Shape3 
         Height          =   4935
         Left            =   1080
         Top             =   4080
         Width           =   6255
      End
      Begin VB.Line Line1 
         X1              =   600
         X2              =   600
         Y1              =   4680
         Y2              =   9000
      End
      Begin VB.Line Line2 
         X1              =   1440
         X2              =   600
         Y1              =   9000
         Y2              =   9000
      End
      Begin VB.Label lblmb 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   7200
         TabIndex        =   29
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label lbladd 
         BackStyle       =   0  'Transparent
         Height          =   615
         Left            =   1680
         TabIndex        =   28
         Top             =   3240
         Width           =   3615
      End
      Begin VB.Label lblnm 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   2040
         TabIndex        =   27
         Top             =   2760
         Width           =   2535
      End
      Begin VB.Label lablep 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   375
         Left            =   6240
         TabIndex        =   26
         Top             =   4200
         Width           =   735
      End
      Begin VB.Shape Shape8 
         Height          =   1335
         Left            =   7320
         Top             =   9000
         Width           =   3015
      End
      Begin VB.Shape Shape7 
         Height          =   4935
         Left            =   6120
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label lableamt 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   255
         Left            =   9000
         TabIndex        =   25
         Top             =   4320
         Width           =   1095
      End
      Begin VB.Label lableqt 
         BackStyle       =   0  'Transparent
         Caption         =   "Quantity"
         Height          =   375
         Left            =   7560
         TabIndex        =   24
         Top             =   4200
         Width           =   855
      End
      Begin VB.Shape Shape6 
         Height          =   615
         Left            =   600
         Top             =   4080
         Width           =   9735
      End
      Begin VB.Label label123 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   255
         Left            =   2040
         TabIndex        =   23
         Top             =   4200
         Width           =   1335
      End
      Begin VB.Label lblsrno 
         BackStyle       =   0  'Transparent
         Caption         =   "Sr."
         Height          =   255
         Left            =   720
         TabIndex        =   22
         Top             =   4200
         Width           =   495
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Moblile no"
         Height          =   255
         Left            =   5640
         TabIndex        =   21
         Top             =   3000
         Width           =   975
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill Date"
         Height          =   375
         Left            =   5640
         TabIndex        =   20
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Name of custmer"
         Height          =   255
         Left            =   600
         TabIndex        =   18
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill No "
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   2280
         Width           =   975
      End
      Begin VB.Shape Shape5 
         Height          =   4935
         Left            =   8520
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Datawadi  Near P.L Deshpande Sinhagad Road                                Pune-411030"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   2055
         Left            =   2880
         TabIndex        =   16
         Top             =   840
         Width           =   6135
      End
      Begin VB.Label Label1 
         BackColor       =   &H000000C0&
         Caption         =   "                    HONDA  MY WINGS      "
         BeginProperty Font 
            Name            =   "Britannic Bold"
            Size            =   27.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1815
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   13575
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add With Accessries"
      Height          =   735
      Left            =   11520
      TabIndex        =   3
      Top             =   6000
      Width           =   1575
   End
   Begin VB.TextBox txtqut 
      Height          =   495
      Left            =   13320
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin VB.ComboBox cmbbike 
      Height          =   315
      Left            =   11760
      TabIndex        =   1
      Text            =   "Select Bile"
      Top             =   3960
      Width           =   1935
   End
   Begin VB.ComboBox cmbcust 
      Height          =   315
      Left            =   11760
      TabIndex        =   0
      Text            =   "Select Custmer"
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Label lblhypo 
      Caption         =   "Label4"
      Height          =   495
      Left            =   6000
      TabIndex        =   13
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label lblnsc 
      Caption         =   "Label3"
      Height          =   495
      Left            =   5880
      TabIndex        =   12
      Top             =   4920
      Width           =   1215
   End
   Begin VB.Label lblntax 
      Caption         =   "Label2"
      Height          =   495
      Left            =   5880
      TabIndex        =   11
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label lblnin 
      Caption         =   "Label1"
      Height          =   495
      Left            =   5640
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label lblstd 
      Caption         =   "Label5"
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label lblsc 
      Caption         =   "Label4"
      Height          =   495
      Left            =   3360
      TabIndex        =   8
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label lbltax 
      Caption         =   "Label3"
      Height          =   495
      Left            =   3360
      TabIndex        =   7
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Label lblin 
      Caption         =   "Label2"
      Height          =   495
      Left            =   3360
      TabIndex        =   6
      Top             =   4200
      Width           =   1215
   End
   Begin VB.Label lblno 
      Caption         =   "Label1"
      Height          =   495
      Left            =   3240
      TabIndex        =   5
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11880
      TabIndex        =   4
      Top             =   5040
      Width           =   855
   End
End
Attribute VB_Name = "frmbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub cmbcust_Click()
Module1.retdata ("select customerphone from customer where custname ='" & cmbcust.Text & "'")
If Not rs1.EOF Or Not rs1.BOF Then
lblnm.Caption = cmbcust.Text
lblmb.Caption = rs1.Fields(0)

End If
Module1.retdata ("select customeraddress from customer where custname ='" & cmbcust.Text & "'")
If Not rs1.EOF Or Not rs1.BOF Then
lbladd.Caption = rs1.Fields(0).Value
End If

End Sub

Private Sub Command1_Click()
If cmbbike.Text <> "Select Bil" And txtqut.Text <> "" Then
Module1.numval ("select bprice from bike where bname='" & cmbbike.Text & "'")

For i = 0 To 5 Step 1
 If lblsr(i).Caption = "" Then
      lblsr(i).Caption = i + 1
      lblname(i).Caption = cmbbike.Text
      lblprice(i).Caption = nm
      lblqty(i).Caption = txtqut.Text
      lblamt(i).Caption = lblprice(i).Caption + Val(lblno.Caption) + Val(lblin.Caption) + Val(lbltax.Caption) + Val(lblsc.Caption) + Val(lblstd.Caption) * lblqty(i).Caption
      Exit For
       End If
       Next
       lblta.Caption = Val(lblamt(0)) + Val(lblamt(1)) + Val(lblamt(2)) + Val(lblamt(3)) + Val(lblamt(4)) + Val(lblamt(5))
       lblgst.Caption = Val(lblta.Caption) * 2 / 100
       lbltot.Caption = Val(lblta.Caption) + Val(lblgst.Caption)
      Else
      MsgBox ("select Bike and enter Quntity")
     
      End If
 
End Sub

Private Sub Command2_Click()
If cmbbike.Text <> "Select Bil" And txtqut.Text <> "" Then
Module1.numval ("select bprice from bike where bname='" & cmbbike.Text & "'")

For i = 0 To 5 Step 1
 If lblsr(i).Caption = "" Then
      lblsr(i).Caption = i + 1
      lblname(i).Caption = cmbbike.Text
      lblprice(i).Caption = nm
      lblqty(i).Caption = txtqut.Text
      lblamt(i).Caption = lblprice(i).Caption + Val(lblnin.Caption) + Val(lblntax.Caption) + Val(lblnsc.Caption) + Val(lblhypo.Caption) * lblqty(i).Caption
      Exit For
       End If
       Next
       lblta.Caption = Val(lblamt(0)) + Val(lblamt(1)) + Val(lblamt(2)) + Val(lblamt(3)) + Val(lblamt(4)) + Val(lblamt(5))
       lblgst.Caption = Val(lblta.Caption) * 2 / 100
       lbltot.Caption = Val(lblta.Caption) + Val(lblgst.Caption)
      Else
      MsgBox ("select Bike and enter Quntity")
     
      End If
End Sub

Private Sub Command3_Click()
'If Text1.Text <> "" And lblnm.Caption <> "" And Lbldate.Caption <> "" And lblname.Item.Caption <> "" And lblqty.Item.Caption <> "" And lblamt.Item.Caption <> "" And lblta.Caption <> "" And lblgst.Caption <> "" And lbltot.Caption <> "" Then
'Module1.inupdel ("insert into bill  values ('" & Text1.Text & "','" & lblnm.Caption & "','" & Lbldate.Caption & "','" & lblname.Count.Caption & "','" & lblqty.Count.Caption & "','" & lblamt.Count.Caption & "','" & lblta.Caption & "','" & lblgst.Caption & "','" & lbltot.Caption & "' )")
'MsgBox " Data save "

End Sub

Private Sub Command4_Click()
Me.Hide
MDIForm1.Show
End Sub

Private Sub Form_Load()
'Module1.numval ("select max(bilno)+1 from bill")
 'Text1.Text = nm
Module1.retdata ("select * from bike")
While Not rs1.EOF
cmbbike.AddItem (rs1.Fields(1).Value)
rs1.MoveNext
Wend
Module1.retdata ("select * from customer")
While Not rs1.EOF
cmbcust.AddItem (rs1.Fields(1).Value)
rs1.MoveNext
Wend
Module1.getconnected
Module1.retdata ("select * from customer")
Module1.retdata ("select * from qotation")
lblno.Caption = rs1.Fields(1)
lblin.Caption = rs1.Fields(2)
lbltax.Caption = rs1.Fields(3)
lblsc.Caption = rs1.Fields(4)
lblstd.Caption = rs1.Fields(5)
Module1.retdata ("select * from qotationwtioutaccessries")
lblnin.Caption = rs1.Fields(1)
lblntax.Caption = rs1.Fields(2)
lblnsc.Caption = rs1.Fields(3)
lblhypo.Caption = rs1.Fields(4)
Lbldate.Caption = Date
Module1.numval ("select max(blno)+1 from bill")
Text1.Text = nm
 'Module1.retdata ("select blno from bill )

End Sub





Private Sub print1_Click()
print1.Visible = False
PrintForm
print1.Visible = True
End Sub

