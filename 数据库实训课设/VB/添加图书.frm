VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form Form8 
   Caption         =   "添加图书"
   ClientHeight    =   5955
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9420
   LinkTopic       =   "Form8"
   ScaleHeight     =   5955
   ScaleWidth      =   9420
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   735
      Left            =   5520
      TabIndex        =   15
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   735
      Left            =   2040
      TabIndex        =   14
      Top             =   4920
      Width           =   1695
   End
   Begin VB.PictureBox Picture1 
      Height          =   3975
      Left            =   360
      ScaleHeight     =   3915
      ScaleWidth      =   7995
      TabIndex        =   0
      Top             =   360
      Width           =   8055
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   5160
         TabIndex        =   16
         Top             =   1200
         Width           =   2415
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   450
         Left            =   6240
         Top             =   3480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   794
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
         RecordSource    =   "Books"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   5160
         TabIndex        =   13
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   5160
         TabIndex        =   12
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   1320
         TabIndex        =   11
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   1320
         TabIndex        =   10
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1320
         TabIndex        =   9
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1320
         TabIndex        =   8
         Top             =   360
         Width           =   6255
      End
      Begin VB.Label Label7 
         Caption         =   "单  价"
         Height          =   495
         Left            =   4200
         TabIndex        =   7
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "类  别"
         Height          =   375
         Left            =   4200
         TabIndex        =   6
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "出版时间"
         Height          =   495
         Left            =   4200
         TabIndex        =   5
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "出版社"
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "ISBN号"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "作  者"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "书  名"
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
cn.Open "provider=SQLOLEDB;Data source=DESKTOP-KLVGTUE;User Id=ydczsq; password='881212';Inital Catalog=bookstore"
'rs.Open "select * from books", cn, adOpenKeyset, adLockOptimistic
rs.Open "select * from books", cn, 1, 3

rs.AddNew
rs.Fields("图书名称") = Text1.Text
rs.Fields("作者") = Text2.Text
rs.Fields("类别") = Text7.Text
rs.Fields("ISBN号") = Text3.Text
rs.Fields("价格") = Text5.Text
rs.Fields("出版社") = Text4.Text
rs.Fields("出版时间") = Text6.Text
If Text1.Text <> "" And Text2.Text <> "" And Text7.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" And Text6.Text <> "" Then
    If Int(Text5.Text) > 10000 Or Int(Text5.Text) < 0 Then
        MsgBox "价格错误！无法添加图书。", vbOKOnly + vbExclamation, "提示"
        Text5.Text = ""
        Exit Sub
    End If
    rs.Update
    Adodc1.Refresh
    Form8.Refresh
    Unload Form8
    MsgBox "图书添加成功！", , "提示"
Else
    MsgBox "信息不得为空！", vbOKOnly + vbExclamation, "提示"
End If
Adodc1.Refresh
Form8.Refresh

End Sub

Private Sub Command2_Click()

Text1.Text = ""
Text2.Text = ""
Text7.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""

End Sub

Private Sub Form_Load()

Text1.Text = ""
Text2.Text = ""
Text7.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""

End Sub

