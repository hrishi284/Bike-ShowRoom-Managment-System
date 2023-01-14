VERSION 5.00
Begin VB.Form frmquo 
   BackColor       =   &H8000000E&
   Caption         =   "Form2"
   ClientHeight    =   8490
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13425
   Icon            =   "qq.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Frame1"
      Height          =   10935
      Left            =   720
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      Begin VB.CommandButton print1 
         Caption         =   "Print"
         Height          =   375
         Left            =   3960
         TabIndex        =   35
         Top             =   10560
         Width           =   1935
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Cancel"
         Height          =   495
         Left            =   7440
         TabIndex        =   33
         Top             =   9960
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Back"
         Height          =   495
         Left            =   5400
         TabIndex        =   32
         Top             =   9960
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "With Accessories"
         Height          =   495
         Left            =   2760
         TabIndex        =   31
         Top             =   9960
         Width           =   2415
      End
      Begin VB.ComboBox cmb2 
         Height          =   315
         Left            =   6960
         TabIndex        =   23
         Text            =   "Combo2"
         Top             =   3240
         Width           =   1815
      End
      Begin VB.ComboBox cmb1 
         Height          =   315
         Left            =   1440
         TabIndex        =   22
         Text            =   "Combo1"
         Top             =   3360
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   1320
         TabIndex        =   21
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   6840
         TabIndex        =   20
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1320
         TabIndex        =   19
         Top             =   2160
         Width           =   3255
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Without Accessories"
         Height          =   495
         Left            =   480
         TabIndex        =   17
         Top             =   9960
         Width           =   2055
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Height          =   495
         Left            =   6960
         TabIndex        =   36
         Top             =   2280
         Width           =   1695
      End
      Begin VB.Label lbltot 
         BackStyle       =   0  'Transparent
         Caption         =   "Label17"
         Height          =   255
         Left            =   7440
         TabIndex        =   34
         Top             =   9360
         Width           =   1335
      End
      Begin VB.Label lblacc 
         BackStyle       =   0  'Transparent
         Caption         =   "Label19"
         Height          =   255
         Left            =   7320
         TabIndex        =   30
         Top             =   6960
         Width           =   1575
      End
      Begin VB.Label lblhpa 
         BackStyle       =   0  'Transparent
         Caption         =   "Label19"
         Height          =   255
         Left            =   7320
         TabIndex        =   29
         Top             =   6480
         Width           =   1335
      End
      Begin VB.Label lblsmc 
         BackStyle       =   0  'Transparent
         Caption         =   "Label19"
         Height          =   375
         Left            =   7320
         TabIndex        =   28
         Top             =   5880
         Width           =   1335
      End
      Begin VB.Label lblrto 
         BackStyle       =   0  'Transparent
         Caption         =   "Label19"
         Height          =   375
         Left            =   7320
         TabIndex        =   27
         Top             =   5280
         Width           =   1335
      End
      Begin VB.Label lblin 
         BackStyle       =   0  'Transparent
         Caption         =   "Label19"
         Height          =   255
         Left            =   7320
         TabIndex        =   26
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label lblsprice 
         BackStyle       =   0  'Transparent
         Caption         =   "Label19"
         Height          =   255
         Left            =   7320
         TabIndex        =   25
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "1 Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5040
         TabIndex        =   24
         Top             =   4800
         Width           =   1575
      End
      Begin VB.Label lblacce 
         BackStyle       =   0  'Transparent
         Caption         =   "Accessories"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   18
         Top             =   6960
         Width           =   2415
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6000
         TabIndex        =   16
         Top             =   9360
         Width           =   855
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "On Road Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   15
         Top             =   9360
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Hypothecation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   14
         Top             =   6480
         Width           =   2415
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Smart card"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   13
         Top             =   5880
         Width           =   2415
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "RTO taxt + Other"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   5280
         Width           =   2415
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Insurance"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Ex-showroom price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Colour"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5640
         TabIndex        =   9
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   8
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   7
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Date "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   5
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Dattawadi  Near P.L Deshpande Sinhagad Road                                Pune-411051"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   2055
         Left            =   2280
         TabIndex        =   3
         Top             =   720
         Width           =   6135
      End
      Begin VB.Image Image1 
         Height          =   1260
         Left            =   240
         Picture         =   "qq.frx":30E8C
         Top             =   120
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H000000C0&
         Height          =   1815
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   13575
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
         Left            =   3240
         TabIndex        =   1
         Top             =   960
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmquo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmb1_Click()
Module1.retdata ("select bprice from bike where bname ='" & cmb1.Text & "'")
If Not rs1.EOF Or Not rs1.BOF Then
lblsprice.Caption = rs1.Fields(0)
End If
Module1.retdata ("select insurance from qotation  ")
If Not rs1.EOF Or Not rs1.BOF Then
lblin.Caption = rs1.Fields(0)
End If
Module1.retdata ("select tax from qotation  ")
If Not rs1.EOF Or Not rs1.BOF Then
lblrto.Caption = rs1.Fields(0)
End If
Module1.retdata ("select scard from qotation  ")
If Not rs1.EOF Or Not rs1.BOF Then
lblsmc.Caption = rs1.Fields(0)
End If
Module1.retdata ("select standf from qotation  ")
If Not rs1.EOF Or Not rs1.BOF Then
lblhpa.Caption = rs1.Fields(0)
End If
Module1.retdata ("select access from qotation  ")
If Not rs1.EOF Or Not rs1.BOF Then
lblacc.Caption = rs1.Fields(0)
End If
lbltot.Caption = Val(lblsprice.Caption) + Val(lblin.Caption) + Val(lblrto.Caption) + Val(lblsmc.Caption) + Val(lblhpa.Caption) + Val(lblacc.Caption)

End Sub

Private Sub Command1_Click()
lblacce.Enabled = True
lblacc.Enabled = True

End Sub

Private Sub Command2_Click()
lblacce.Enabled = False
lblacc.Enabled = False

End Sub

Private Sub Command3_Click()
Me.Hide
MDIForm1.Show
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Form_Load()
Label17.Caption = Date
Module1.retdata ("select * from bike")
While Not rs1.EOF
cmb1.AddItem (rs1.Fields(1).Value)
rs1.MoveNext
Wend
Module1.retdata ("select * from bike")
While Not rs1.EOF
cmb2.AddItem (rs1.Fields(4).Value)
rs1.MoveNext
Wend
End Sub



Private Sub print1_Click()
print1.Visible = False
Print
print1.Visible = True
End Sub
