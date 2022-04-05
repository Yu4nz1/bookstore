VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   Caption         =   "用户个人信息"
   ClientHeight    =   7395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14685
   LinkTopic       =   "Form5"
   ScaleHeight     =   7395
   ScaleWidth      =   14685
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command3 
      Caption         =   "下一位"
      Height          =   735
      Left            =   6480
      TabIndex        =   13
      Top             =   6240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "上一位"
      Height          =   615
      Left            =   3120
      TabIndex        =   12
      Top             =   6240
      Width           =   1575
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
      Height          =   5775
      Left            =   600
      ScaleHeight     =   5715
      ScaleWidth      =   12195
      TabIndex        =   1
      Top             =   120
      Width           =   12255
      Begin VB.PictureBox Picture2 
         Height          =   2775
         Left            =   8400
         ScaleHeight     =   2715
         ScaleWidth      =   2355
         TabIndex        =   14
         Top             =   840
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         DataField       =   "用户名"
         DataSource      =   "Adodc1"
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
         TabIndex        =   6
         Top             =   480
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         DataField       =   "密码"
         DataSource      =   "Adodc1"
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
         TabIndex        =   5
         Top             =   1560
         Width           =   3375
      End
      Begin VB.TextBox Text4 
         DataField       =   "姓名"
         DataSource      =   "Adodc1"
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
         TabIndex        =   4
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox Text5 
         DataField       =   "联系电话"
         DataSource      =   "Adodc1"
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
         TabIndex        =   3
         Top             =   3720
         Width           =   3375
      End
      Begin VB.TextBox Text6 
         DataField       =   "地址"
         DataSource      =   "Adodc1"
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
         TabIndex        =   2
         Top             =   4800
         Width           =   7815
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   735
         Left            =   8640
         Top             =   3840
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
         Connect         =   "DSN=书本信息"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "书本信息"
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
         TabIndex        =   11
         Top             =   600
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
         TabIndex        =   10
         Top             =   1560
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
         TabIndex        =   9
         Top             =   2640
         Width           =   1575
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
         TabIndex        =   8
         Top             =   3720
         Width           =   1575
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
         TabIndex        =   7
         Top             =   4800
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "退出"
      Height          =   735
      Left            =   10200
      TabIndex        =   0
      Top             =   6240
      Width           =   1935
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Enabled = True
Text2.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then Adodc1.Recordset.MoveFirst
xh = Adodc1.Recordset.Fields("用户名")
X1 = App.Path & "\user\" & xh & ".bmp"
Picture1.Picture = LoadPicture(X1)

End Sub

Private Sub Command2_Click()

Load Form4
Form4.Show
End Sub

Private Sub Command3_Click()
Text1.Enabled = True
Text2.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast
xh = Adodc1.Recordset.Fields("用户名")
X1 = App.Path & "\user\" & xh & "    .bmp"
Picture1.Picture = LoadPicture(X1)

End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
End Sub
