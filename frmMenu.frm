VERSION 5.00
Begin VB.Form frmMenu 
   ClientHeight    =   30
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   1560
   LinkTopic       =   "Form2"
   ScaleHeight     =   30
   ScaleWidth      =   1560
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuzoom 
      Caption         =   "mnuZoom"
      Begin VB.Menu mnu1 
         Caption         =   "10%"
      End
      Begin VB.Menu mnu2 
         Caption         =   "20%"
      End
      Begin VB.Menu mnu3 
         Caption         =   "30%"
      End
      Begin VB.Menu mnu4 
         Caption         =   "40%"
      End
      Begin VB.Menu mnu5 
         Caption         =   "50%"
      End
      Begin VB.Menu mnu6 
         Caption         =   "60%"
      End
      Begin VB.Menu mnu7 
         Caption         =   "70%"
      End
      Begin VB.Menu mnu8 
         Caption         =   "80%"
      End
      Begin VB.Menu mnu9 
         Caption         =   "90%"
      End
      Begin VB.Menu mnu10 
         Caption         =   "100%"
      End
      Begin VB.Menu mnu12 
         Caption         =   "120%"
      End
      Begin VB.Menu mnu14 
         Caption         =   "140%"
      End
      Begin VB.Menu mnu16 
         Caption         =   "160%"
      End
      Begin VB.Menu mnu18 
         Caption         =   "180%"
      End
      Begin VB.Menu mnu20 
         Caption         =   "200%"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnu1_Click()
    ZoomFactor = 10
End Sub

Private Sub mnu10_Click()
    ZoomFactor = 100
End Sub

Private Sub mnu12_Click()
    ZoomFactor = 120
End Sub

Private Sub mnu14_Click()
    ZoomFactor = 140
End Sub

Private Sub mnu16_Click()
    ZoomFactor = 160
End Sub

Private Sub mnu18_Click()
    ZoomFactor = 180
End Sub

Private Sub mnu2_Click()
    ZoomFactor = 20
End Sub

Private Sub mnu20_Click()
    ZoomFactor = 200
End Sub

Private Sub mnu3_Click()
    ZoomFactor = 30
End Sub

Private Sub mnu4_Click()
    ZoomFactor = 40
End Sub

Private Sub mnu5_Click()
    ZoomFactor = 50
End Sub

Private Sub mnu6_Click()
    ZoomFactor = 60
End Sub

Private Sub mnu7_Click()
    ZoomFactor = 70
End Sub

Private Sub mnu8_Click()
    ZoomFactor = 80
End Sub

Private Sub mnu9_Click()
    ZoomFactor = 90
End Sub
