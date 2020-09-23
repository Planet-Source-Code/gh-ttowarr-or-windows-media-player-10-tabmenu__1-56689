VERSION 5.00
Begin VB.UserControl TabButtonMenu 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin Project1.TabSeparator TabSeparator 
      Height          =   375
      Left            =   330
      TabIndex        =   0
      Top             =   100
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   661
   End
   Begin VB.Image ImgDown 
      Height          =   540
      Left            =   2520
      Picture         =   "TabButtonMenu.ctx":0000
      Top             =   960
      Width           =   330
   End
   Begin VB.Image ImgNormal 
      Height          =   540
      Left            =   2880
      Picture         =   "TabButtonMenu.ctx":09D2
      Top             =   960
      Width           =   330
   End
   Begin VB.Image ImgOver 
      Height          =   540
      Left            =   2160
      Picture         =   "TabButtonMenu.ctx":13A4
      Top             =   960
      Width           =   330
   End
   Begin VB.Image ImgButton 
      Height          =   525
      Left            =   0
      Top             =   0
      Width           =   330
   End
   Begin VB.Image ImgBG 
      Height          =   525
      Left            =   0
      Picture         =   "TabButtonMenu.ctx":1D76
      Stretch         =   -1  'True
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "TabButtonMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event Click()
Event MouseMove()

Private Sub ImgButton_Click()
Click
End Sub

Private Sub ImgButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDown
End Sub

Private Sub ImgButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMove
End Sub

Private Sub UserControl_Initialize()
ImgButton.Picture = ImgNormal.Picture
End Sub

Private Sub MouseMove()
RaiseEvent MouseMove
ImgButton.Picture = ImgOver.Picture
End Sub

Private Sub Click()
RaiseEvent Click
End Sub

Private Sub MouseDown()
ImgButton.Picture = ImgDown.Picture
End Sub

Public Sub MoveMouseOff()
ImgButton.Picture = ImgNormal.Picture
End Sub

Private Sub UserControl_Resize()
UserControl.Height = 525
UserControl.Width = 405
ImgBG.Width = UserControl.Width
ImgBG.Top = 0
End Sub
