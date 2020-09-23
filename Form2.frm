VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3240
   ClientLeft      =   165
   ClientTop       =   645
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3240
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnupopup 
      Caption         =   "&File"
      Begin VB.Menu mnuClear 
         Caption         =   "Clear Playlist"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mnuClear_Click()
Form1.PlayList.Clear
Form1.List1.Clear
End Sub
