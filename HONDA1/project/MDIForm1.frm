VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   9300
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15120
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":30E8C
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu Del 
      Caption         =   "Master Details "
      Begin VB.Menu Emp 
         Caption         =   "Employe Details"
         Shortcut        =   ^E
      End
      Begin VB.Menu cus 
         Caption         =   "Custmer Details "
         Shortcut        =   ^J
      End
      Begin VB.Menu BikeD 
         Caption         =   "Bike Details"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu search 
      Caption         =   "Search Master"
      Begin VB.Menu custmer 
         Caption         =   "Custmer"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Employee 
         Caption         =   "Employee"
         Shortcut        =   {F2}
      End
      Begin VB.Menu Bike 
         Caption         =   "Bike"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu frmupd 
      Caption         =   "Update Detais"
      Begin VB.Menu ubik 
         Caption         =   "Bike"
      End
      Begin VB.Menu ucm 
         Caption         =   "Custmer"
      End
      Begin VB.Menu uem 
         Caption         =   "Employee"
      End
   End
   Begin VB.Menu Enquiry 
      Caption         =   "Enquiry"
   End
   Begin VB.Menu Quotation 
      Caption         =   "Quotation"
   End
   Begin VB.Menu emi 
      Caption         =   "EMI"
   End
   Begin VB.Menu bill 
      Caption         =   "Billing"
   End
   Begin VB.Menu rep 
      Caption         =   "Report"
      Begin VB.Menu brep 
         Caption         =   "Bike Report "
      End
      Begin VB.Menu er 
         Caption         =   "Employe Report"
      End
      Begin VB.Menu CR 
         Caption         =   "Custmer Report"
      End
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cust_Click()

End Sub

Private Sub Bike_Click()
Me.Hide
frmsbike.Show
End Sub

Private Sub BikeD_Click()
Me.Hide
frmbike.Show
End Sub

Private Sub bill_Click()
Me.Hide
frmbill.Show

End Sub

Private Sub brep_Click()
Me.Hide
DataReport1.Show
End Sub

Private Sub CR_Click()
Me.Hide
DataReport3.Show
End Sub

Private Sub Cus_Click()
Me.Hide
frmCust.Show
End Sub

Private Sub custmer_Click()
Me.Hide
frmscust.Show
End Sub

Private Sub emi_Click()
Me.Hide
frmEmi.Show
End Sub

Private Sub Emp_Click()
Me.Hide
frmEmp.Show
End Sub

Private Sub Employee_Click()
Me.Hide
frmsEmp.Show
End Sub

Private Sub Enquiry_Click()
Me.Hide
frmenq.Show
End Sub

Private Sub er_Click()
Me.Hide
DataReport2.Show

End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Quotation_Click()
Me.Hide
frmquo.Show
End Sub


Private Sub ubik_Click()
Me.Hide
frmubk.Show
End Sub

Private Sub ucm_Click()
Me.Hide
frmuc.Show

End Sub

Private Sub uem_Click()
Me.Hide
frmuep.Show

End Sub
