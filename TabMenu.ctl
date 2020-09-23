VERSION 5.00
Begin VB.UserControl TabMenu 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin Project1.TabButton TabButton 
      Height          =   525
      Index           =   1
      Left            =   840
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   926
   End
   Begin Project1.TabButtonMenu TabButtonMenu 
      Height          =   525
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   926
   End
   Begin VB.Image ImgBG 
      Height          =   525
      Left            =   0
      Picture         =   "TabMenu.ctx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "TabMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event ItemClick()
Event ItemMenuClick()
Public ItemText As String

Private Sub TabButton_Click(Index As Integer)
RemoveTabFocus Index
ItemText = TabButton(Index).Text
RaiseEvent ItemClick
End Sub

Private Sub TabButton_MouseMove(Index As Integer)
MoveMouseOff Index
End Sub

Private Sub TabButtonMenu_Click(Index As Integer)
ItemText = Index - 1
RaiseEvent ItemMenuClick
End Sub

Private Sub TabButtonMenu_MouseMove(Index As Integer)
MoveMouseOff Index
End Sub

Private Sub UserControl_Resize()
ImgBG.Width = UserControl.Width
UserControl.Height = ImgBG.Height
End Sub

Public Sub AddButton(Style_1_Tab_2_Menu As Integer, Optional ItemText As String)
Dim i  As Integer
Dim a As Integer
If Style_1_Tab_2_Menu = 1 Then
i = TabButton.Count + 1
Load TabButton(i)
TabButton(i).Top = 0
TabButton(i).Left = (TabButton.Count * 1050 - 1050 - 1050) + (TabButtonMenu.Count * 405 - 405)
TabButton(i).SetText ItemText
TabButton(i).Visible = True
ElseIf Style_1_Tab_2_Menu = 2 Then
a = TabButtonMenu.Count + 1
Load TabButtonMenu(a)
TabButtonMenu(a).Top = 0
TabButtonMenu(a).Left = (TabButton.Count * 1050 - 1050) + (TabButtonMenu.Count * 405 - 405 - 405)
TabButtonMenu(a).Visible = True
End If
End Sub

Public Sub MoveMouseOff(Optional Index As Integer)
If Index = 0 Then Index = 0
Dim i As Integer
For i = 1 To TabButton.Count
If Not i = Index Then
TabButton(i).MoveMouseOff
End If
Next i
For i = 1 To TabButtonMenu.Count
If Not i = Index Then
TabButtonMenu(i).MoveMouseOff
End If
Next i
End Sub

Public Sub RemoveTabFocus(Optional Index As Integer)
If Index = 0 Then Index = 0
Dim i As Integer
For i = 1 To TabButton.Count
If Not i = Index Then
TabButton(i).RemoveFocus
End If
Next i
End Sub

Public Sub SetTabFocus(Index As Integer)
TabButton(Index + 1).Click
End Sub
