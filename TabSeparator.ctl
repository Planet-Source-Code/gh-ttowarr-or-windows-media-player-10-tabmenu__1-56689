VERSION 5.00
Begin VB.UserControl TabSeparator 
   ClientHeight    =   1095
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1230
   ScaleHeight     =   1095
   ScaleWidth      =   1230
   Begin VB.Image ImgSep 
      Height          =   375
      Left            =   0
      Picture         =   "TabSeparator.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   70
   End
End
Attribute VB_Name = "TabSeparator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub UserControl_Resize()
UserControl.Height = ImgSep.Height
UserControl.Width = ImgSep.Width
End Sub
