VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "网上书店管理系统"
   ClientHeight    =   11415
   ClientLeft      =   225
   ClientTop       =   915
   ClientWidth     =   18390
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11415
   ScaleWidth      =   18390
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Height          =   5775
      Left            =   5640
      Picture         =   "网上书店管理系统.frx":0000
      ScaleHeight     =   5715
      ScaleWidth      =   7155
      TabIndex        =   0
      Top             =   2880
      Width           =   7215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "     BOOKSTORE"
      BeginProperty Font 
         Name            =   "SimSun-ExtB"
         Size            =   42
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   8775
   End
   Begin VB.Menu 网上书店管理系统 
      Caption         =   "网上书店管理系统"
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu 网上 
      Caption         =   "网上"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu 用户 
      Caption         =   "用户"
      NegotiatePosition=   3  'Right
      Begin VB.Menu 用户注册 
         Caption         =   "用户注册"
      End
      Begin VB.Menu 用户登录 
         Caption         =   "用户登录"
      End
   End
   Begin VB.Menu 书店 
      Caption         =   "书店"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu 注册 
      Caption         =   "管理员"
      NegotiatePosition=   3  'Right
      Begin VB.Menu 管理员登录 
         Caption         =   "管理员登录"
      End
   End
   Begin VB.Menu 退出 
      Caption         =   "退出"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()

conn = "provider = SQLOLEDB ; Data source = DESKTOP-KLVGTUE ; USER ID = ydczsq ; password = '881212' ; inital catalog = Bookstore_Management"
'cn.Open conn

End Sub

Private Sub 管理员登录_Click()

Load login
login.Show

End Sub


Private Sub 退出_Click()
sign_out = MsgBox("是否退出系统(Y/N)？", vbYesNo, "提示")
If sign_out = 6 Then
    For i = 0 To Forms.count - 1
        Unload Forms(0)
    Next i
End If
End Sub

Private Sub 用户登录_Click()

Load login
login.Show

End Sub

Private Sub 用户注册_Click()

Load register
register.Show


End Sub
