VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "����������ϵͳ"
   ClientHeight    =   11415
   ClientLeft      =   225
   ClientTop       =   915
   ClientWidth     =   18390
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   11415
   ScaleWidth      =   18390
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox Picture1 
      Height          =   5775
      Left            =   5640
      Picture         =   "����������ϵͳ.frx":0000
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
   Begin VB.Menu ����������ϵͳ 
      Caption         =   "����������ϵͳ"
      NegotiatePosition=   1  'Left
   End
   Begin VB.Menu ���� 
      Caption         =   "����"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu �û� 
      Caption         =   "�û�"
      NegotiatePosition=   3  'Right
      Begin VB.Menu �û�ע�� 
         Caption         =   "�û�ע��"
      End
      Begin VB.Menu �û���¼ 
         Caption         =   "�û���¼"
      End
   End
   Begin VB.Menu ��� 
      Caption         =   "���"
      Enabled         =   0   'False
      Visible         =   0   'False
   End
   Begin VB.Menu ע�� 
      Caption         =   "����Ա"
      NegotiatePosition=   3  'Right
      Begin VB.Menu ����Ա��¼ 
         Caption         =   "����Ա��¼"
      End
   End
   Begin VB.Menu �˳� 
      Caption         =   "�˳�"
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

Private Sub ����Ա��¼_Click()

Load login
login.Show

End Sub


Private Sub �˳�_Click()
sign_out = MsgBox("�Ƿ��˳�ϵͳ(Y/N)��", vbYesNo, "��ʾ")
If sign_out = 6 Then
    For i = 0 To Forms.count - 1
        Unload Forms(0)
    Next i
End If
End Sub

Private Sub �û���¼_Click()

Load login
login.Show

End Sub

Private Sub �û�ע��_Click()

Load register
register.Show


End Sub
