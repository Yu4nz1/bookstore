VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form register 
   Caption         =   "注册"
   ClientHeight    =   8505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13800
   LinkTopic       =   "Form3"
   ScaleHeight     =   8505
   ScaleWidth      =   13800
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      TabIndex        =   2
      Top             =   7080
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2040
      TabIndex        =   1
      Top             =   7080
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   960
      ScaleHeight     =   5595
      ScaleWidth      =   11715
      TabIndex        =   0
      Top             =   720
      Width           =   11775
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   735
         Left            =   9000
         Top             =   3360
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   1296
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
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2640
         TabIndex        =   14
         Top             =   4680
         Width           =   7815
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   13
         Top             =   3840
         Width           =   3375
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   12
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   11
         Top             =   2160
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   10
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2640
         TabIndex        =   9
         Top             =   480
         Width           =   3375
      End
      Begin VB.Label Label6 
         Caption         =   "地  址："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   8
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "电  话："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   7
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "姓  名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   6
         Top             =   3120
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "确认密码："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   5
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "密  码："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   4
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "用户名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
   End
End
Attribute VB_Name = "register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
cn.Open "provider=SQLOLEDB;Data source=DESKTOP-KLVGTUE;User Id=ydczsq; password='881212';Inital Catalog=bookstore"
rs.Open "select * from Users", cn
If Text1.Text = "" Or Text2.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then
    MsgBox "信息不得为空!", vbOKOnly + vbExclamation, "提示"
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
End If

Do While Not rs.EOF
    If Trim(Text1.Text) = Trim(rs.Fields("用户名")) Then Exit Do '用户名已存在
    rs.MoveNext
Loop

If Not rs.EOF Then
    MsgBox " '" & Text1.Text & " '该用户名已存在！", vbOKOnly + vbExclamation, "提示"
    Text1.Text = ""
End If

If rs.EOF And Text1.Text <> "" Then
    cn.Execute "insert into users(用户名,密码,姓名,地址,联系电话,权限)values(' " & Text1.Text & " ',' " & Text2.Text & " ',' " & Text4.Text & " ',' " & Text6.Text & " ',' " & Text5.Text & " ',0)"
    MsgBox " '" & Text1.Text & " '注册成功!", vbOKOnly + vbExclamation, "提示"
    Adodc1.Refresh
    Unload register
    Load login
    login.Show
    
End If

Adodc1.Refresh
register.Refresh

End Sub

Private Sub Command2_Click()
'取消
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""

End Sub


Private Sub Form_Load()

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""

End Sub


Private Sub Text2_Click()

Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
cn.Open "provider=SQLOLEDB;Data source=DESKTOP-KLVGTUE;User Id=ydczsq; password='881212';Inital Catalog=bookstore"
rs.Open "select * from Users", cn
Do While Not rs.EOF
    If Trim(Text1.Text) = Trim(rs.Fields("用户名")) Then Exit Do '用户名已存在
    rs.MoveNext
Loop
If Not rs.EOF Then
    MsgBox " '" & Text1.Text & " '该用户名已存在！", vbOKOnly + vbExclamation, "提示"
    Text1.Text = ""
End If
End Sub

Private Sub Text4_Click()
If Trim(Text2.Text) <> Trim(Text3.Text) Then
    MsgBox "两次密码不一致！", vbOKOnly + vbExclamation, "提示"
    Text2.Text = ""
    Text3.Text = ""
End If
End Sub

