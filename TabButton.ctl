VERSION 5.00
Begin VB.UserControl TabButton 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   ScaleHeight     =   3600
   ScaleWidth      =   5400
   Begin Project1.TabSeparator TabSeparator 
      Height          =   375
      Left            =   975
      TabIndex        =   0
      Top             =   100
      Width           =   75
      _ExtentX        =   132
      _ExtentY        =   661
   End
   Begin VB.Label LblText 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ItemText"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   190
      Width           =   975
   End
   Begin VB.Image ImgButton 
      Height          =   525
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   975
   End
   Begin VB.Image ImgOver 
      Height          =   540
      Left            =   3600
      Picture         =   "TabButton.ctx":0000
      Top             =   120
      Width           =   960
   End
   Begin VB.Image ImgNormal 
      Height          =   540
      Left            =   2520
      Picture         =   "TabButton.ctx":1B42
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Image ImgDown 
      Height          =   525
      Left            =   1440
      Picture         =   "TabButton.ctx":1C9C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   975
   End
   Begin VB.Image ImgBG 
      Height          =   525
      Left            =   0
      Picture         =   "TabButton.ctx":386E
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "TabButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event Click()
Event MouseMove()
Public Focused As Boolean
Public Text As String

Private Sub ImgButton_Click()
Click
End Sub

Private Sub ImgButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDown
End Sub

Private Sub ImgButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMove
End Sub

Private Sub LblText_Click()
Click
End Sub

Private Sub LblText_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseDown
End Sub

Private Sub LblText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseMove
End Sub

Private Sub UserControl_Initialize()
ImgButton.Picture = ImgNormal.Picture
End Sub

Private Sub MouseMove()
RaiseEvent MouseMove
If Focused = False Then
ImgButton.Picture = ImgOver.Picture
End If
End Sub

Public Sub Click()
RaiseEvent Click
SetFocus
End Sub

Private Sub MouseDown()
ImgButton.Picture = ImgDown.Picture
End Sub

Public Sub MoveMouseOff()
If Focused = False Then
ImgButton.Picture = ImgNormal.Picture
End If
End Sub

Private Sub UserControl_Resize()
UserControl.Height = 525
UserControl.Width = 1050
ImgBG.Width = UserControl.Width
ImgBG.Top = 0
End Sub

Public Sub SetFocus()
Focused = True
ImgButton.Picture = ImgDown.Picture
LblText.ForeColor = &HFFFFFF
End Sub

Public Sub RemoveFocus()
Focused = False
ImgButton.Picture = ImgNormal.Picture
LblText.ForeColor = &H400000
End Sub

Public Sub SetText(ItemText As String)
LblText.Caption = ItemText
Text = ItemText
End Sub
