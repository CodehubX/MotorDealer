VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "MDIForm1"
   ClientHeight    =   5130
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   5265
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu hm 
      Caption         =   "Home"
      Begin VB.Menu lgt 
         Caption         =   "Logout"
      End
   End
   Begin VB.Menu cst 
      Caption         =   "Customer"
      Begin VB.Menu pbl 
         Caption         =   "Pembeli"
      End
   End
   Begin VB.Menu prd 
      Caption         =   "Produc"
      Begin VB.Menu mtr 
         Caption         =   "Motor"
      End
   End
   Begin VB.Menu by 
      Caption         =   "Buy"
      Begin VB.Menu chs 
         Caption         =   "Chas"
      End
      Begin VB.Menu krd 
         Caption         =   "Kredit"
      End
   End
   Begin VB.Menu rpt 
      Caption         =   "Report"
      Begin VB.Menu pp 
         Caption         =   "Print Pelanggan"
      End
      Begin VB.Menu pm 
         Caption         =   "Print Motor"
      End
      Begin VB.Menu pc 
         Caption         =   "Print Chas"
      End
      Begin VB.Menu pk 
         Caption         =   "Print Kredit"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chs_Click()
chas.Show
Me.Hide
End Sub

Private Sub krd_Click()
credit.Show
Me.Hide
End Sub

Private Sub lgt_Click()
login.Show
Me.Hide
End Sub

Private Sub mtr_Click()
motor.Show
Me.Hide
End Sub

Private Sub pbl_Click()
pembeli.Show
Me.Hide
End Sub
