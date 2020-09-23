VERSION 5.00
Begin VB.MDIForm mdiIntersection 
   BackColor       =   &H8000000C&
   Caption         =   "My Intersection"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu File 
      Caption         =   "&File"
      Begin VB.Menu NewGame 
         Caption         =   "&New Game"
      End
      Begin VB.Menu Save 
         Caption         =   "&Save And Exit"
      End
      Begin VB.Menu Load 
         Caption         =   "&Load Game"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "mdiIntersection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Exit_Click()
    frmIntersection.cmdCode_Click
    
End Sub

Private Sub Load_Click()
    frmIntersection.cmdLoad_Click
End Sub

Private Sub NewGame_Click()
    frmIntersection.cmdNewGame_Click
End Sub

Private Sub Save_Click()
    frmIntersection.cmdSave_Click
End Sub

