VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form login 
   Caption         =   "登录"
   ClientHeight    =   6015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15.75
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "登录"
   ScaleHeight     =   6015
   ScaleWidth      =   7575
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   735
      Left            =   4200
      TabIndex        =   2
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   735
      Left            =   1320
      TabIndex        =   1
      Top             =   4680
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   600
      ScaleHeight     =   3435
      ScaleWidth      =   6315
      TabIndex        =   0
      Top             =   720
      Width           =   6375
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   4080
         Top             =   2880
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1085
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=1"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "1"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Users"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox Text2 
         Height          =   675
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   1920
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   735
         Left            =   1800
         TabIndex        =   5
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "密  码："
         Height          =   735
         Left            =   480
         TabIndex        =   4
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "用户名："
         Height          =   735
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'登录
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
cn.Open "provider=SQLOLEDB;Data source=DESKTOP-KLVGTUE;User Id=ydczsq; password='881212';Inital Catalog=bookstore"
rs.Open "select * from Users", cn
Do While Not rs.EOF
    If Trim(Text1.Text) = Trim(rs.Fields("用户名")) And Trim(Text2.Text) = Trim(rs.Fields("密码")) Then
        If rs.Fields("权限") = "0" Then
            'Public username As String
            username = Trim(Text1.Text)
           ' MsgBox " '" & username & " '登录成功!", , "提示"
            Form4.Show
        Else
            Form7.Show
        End If
        Unload login
        Exit Do
    End If
    rs.MoveNext
Loop
If rs.EOF Then
    MsgBox "用户名或密码错误！", vbOKOnly + vbExclamation, "提示"
End If
Adodc1.Refresh
Form1.Refresh

End Sub

Private Sub Command2_Click()

Text1.Text = ""
Text2.Text = ""

End Sub

Private Sub Form_Load()

Text1.Text = ""
Text2.Text = ""


End Sub

