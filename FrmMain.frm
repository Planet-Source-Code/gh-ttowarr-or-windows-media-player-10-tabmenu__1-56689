VERSION 5.00
Begin VB.Form FrmMain 
   ClientHeight    =   6090
   ClientLeft      =   4065
   ClientTop       =   4335
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8535
   Begin VB.CommandButton CmdInfo 
      Caption         =   "Info"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox TxtFocus 
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Text            =   "2"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton CmdSetFocus 
      Caption         =   "Set Focus"
      Height          =   375
      Left            =   3720
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton CmdAddMenu 
      Caption         =   "Add MenuButton"
      Height          =   375
      Left            =   5280
      TabIndex        =   3
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox TxtText 
      Height          =   375
      Left            =   6840
      TabIndex        =   2
      Text            =   "ButtonText"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add Button"
      Height          =   375
      Left            =   6840
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin Project1.TabMenu TabMenu 
      Height          =   525
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   926
   End
   Begin VB.Image ImgBG 
      Height          =   1290
      Left            =   0
      Picture         =   "FrmMain.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1410
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAdd_Click()
TabMenu.AddButton 1, TxtText.Text
End Sub

Private Sub CmdAddMenu_Click()
TabMenu.AddButton 2
End Sub

Private Sub CmdInfo_Click()
MsgBox "You can only focus the buttons not the buttonmenu's"
End Sub

Private Sub CmdSetFocus_Click()
TabMenu.SetTabFocus TxtFocus.Text
End Sub

Private Sub Form_Load()
TabMenu.AddButton 1, "About"
TabMenu.AddButton 1, "DemoTab1"
TabMenu.AddButton 1, "DemoTab2"
TabMenu.AddButton 1, "DemoTab3"
TabMenu.AddButton 2
TabMenu.SetTabFocus 2
End Sub

Private Sub Form_Resize()
TabMenu.Width = Me.Width
ImgBG.Width = Me.Width
ImgBG.Height = Me.Height
End Sub

Private Sub ImgBG_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
TabMenu.MoveMouseOff
End Sub

Private Sub TabMenu_ItemClick()
If TabMenu.ItemText = "About" Then MsgBox "Made by Gh€ttoWarr!or", vbInformation, "About"
End Sub

Private Sub TabMenu_ItemMenuClick()
MsgBox "This is a PopUpMenuButton", vbInformation, "PopUpMenuButton"
End Sub
