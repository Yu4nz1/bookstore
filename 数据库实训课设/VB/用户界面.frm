VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form4 
   Caption         =   "用户界面"
   ClientHeight    =   9060
   ClientLeft      =   225
   ClientTop       =   915
   ClientWidth     =   17295
   LinkTopic       =   "Form4"
   ScaleHeight     =   9060
   ScaleWidth      =   17295
   StartUpPosition =   3  '窗口缺省
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1800
      Top             =   8160
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   582
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
      RecordSource    =   "shopping"
      Caption         =   "Adodc2"
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
   Begin VB.CommandButton Command6 
      Caption         =   "下一本"
      Height          =   735
      Left            =   12240
      TabIndex        =   14
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "上一本"
      Height          =   735
      Left            =   9840
      TabIndex        =   13
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "加入订单"
      Height          =   735
      Left            =   14400
      TabIndex        =   1
      Top             =   7920
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   6855
      Left            =   960
      ScaleHeight     =   6795
      ScaleWidth      =   15795
      TabIndex        =   0
      Top             =   960
      Width           =   15855
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "用户界面.frx":0000
         Height          =   3255
         Left            =   600
         TabIndex        =   16
         Top             =   3240
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Caption         =   "图书浏览"
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   1680
         TabIndex        =   15
         Top             =   1320
         Width           =   2415
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   840
         Top             =   2160
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
      Begin VB.CommandButton Command1 
         Caption         =   "取消"
         Height          =   735
         Left            =   12360
         TabIndex        =   12
         Top             =   2040
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "查询"
         Height          =   735
         Left            =   8280
         TabIndex        =   11
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   10200
         TabIndex        =   10
         Top             =   1320
         Width           =   4815
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   5760
         TabIndex        =   9
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   10200
         TabIndex        =   8
         Top             =   480
         Width           =   4815
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1680
         TabIndex        =   7
         Top             =   480
         Width           =   6375
      End
      Begin VB.CheckBox Check5 
         Caption         =   "出版社"
         Height          =   420
         Left            =   9000
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox Check4 
         Caption         =   "ISBN号"
         Height          =   495
         Left            =   9000
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox Check3 
         Caption         =   "作  者"
         Height          =   420
         Left            =   4560
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "分  类"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "书  名"
         Height          =   420
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Menu 我 
      Caption         =   "我"
      Index           =   0
   End
   Begin VB.Menu 个人信息 
      Caption         =   "个人信息"
   End
   Begin VB.Menu 我的订单 
      Caption         =   "我的订单"
   End
   Begin VB.Menu 退出 
      Caption         =   "退出"
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'取消
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Connection
Dim sql As String
cn.Open "provider=SQLOLEDB;Data source=DESKTOP-KLVGTUE;User Id=ydczsq; password='881212';Inital Catalog=bookstore"
Adodc1.ConnectionString = "dsn=1"
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from Books "
Adodc1.Refresh
'yn2 = MsgBox("确实退出系统吗(Y/N)?", vbYesNo, "退出系统提示")
'If yn2 = 6 Then
'For i = 0 To Forms.count - 1
'  Unload Forms(0)
'Next i
'End If
End Sub

Private Sub Command3_Click()
'加入订单
Dim bookname, phonenumber, adress As String
Dim prices As Integer
bookname = Adodc1.Recordset("图书名称")
prices = Adodc1.Recordset("价格")
'username
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
cn.Open "provider=SQLOLEDB;Data source=DESKTOP-KLVGTUE;User Id=ydczsq; password='881212';Inital Catalog=Bookstore"
rs.Open "select * from users", cn
Do While Not rs.EOF
    If Trim(rs.Fields("用户名")) = username Then
        phonenumber = rs.Fields("联系电话")
        adress = rs.Fields("地址")
        Exit Do
    End If
    rs.MoveNext
Loop
rs.Close
cn.Execute "insert into shopping(图书名称,价格,用户姓名,发货日期,状态,电话,地址) values ('" & bookname & "'," & prices & ",'" & username & "','','','" & phonenumber & "','" & adress & "')"
Adodc2.Refresh
'Dim cn As New ADODB.Connection
'Dim rs As New ADODB.Recordset
'cn.Open "provider=SQLOLEDB;Data source=DESKTOP-KLVGTUE;User Id=ydczsq; password='881212';Inital Catalog=Bookstore"
'cn.Execute "insert into shopping(图书名称,价格) values('" & Text1.Text & "','" & Text2.Text & " ','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "')"
'Adodc2.ConnectionString = "dsn=1"
'Adodc2.CommandType = adCmdUnknown
'Adodc2.RecordSource = ""
'Form1.Refresh

'Adodc2.Refresh

End Sub

Private Sub Command4_Click()
'查询
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Connection
Dim sql As String
cn.Open "provider=SQLOLEDB;Data source=DESKTOP-KLVGTUE;User Id=ydczsq; password='881212';Inital Catalog=bookstore"
Adodc1.ConnectionString = "dsn=1"
Adodc1.CommandType = adCmdUnknown
'sql = "select * from Books "
If Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" And Text4.Text <> "" And Text5.Text <> "" Then
    Adodc1.RecordSource = "select * from Books where ISBN号= '" & Text2.Text & " 'and 图书名称='" & Text1.Text & " 'and 出版社='" & Text4.Text & " 'and 作者='" & Text3.Text & " 'and 类别='" & Text5.Text & " '"
ElseIf Text1.Text <> "" Then
    Adodc1.RecordSource = "select * from Books where 图书名称='" & Text1.Text & " '"
ElseIf Text2.Text <> "" Then
    Adodc1.RecordSource = "select * from Books where ISBN号= '" & Text2.Text & " '"
ElseIf Text4.Text <> "" Then
    Adodc1.RecordSource = "select * from Books where 出版社='" & Text4.Text & " '"
ElseIf Text3.Text <> "" Then
    Adodc1.RecordSource = "select * from Books where 作者='" & Text3.Text & " '"
ElseIf Text5.Text <> "" Then
    Adodc1.RecordSource = "select * from Books where 类别='" & Text5.Text & " '"
Else
    MsgBox "无法查询，请重新输入所需信息！", , "提示"
    Adodc1.RecordSource = "select * from Books "
End If
Adodc1.Refresh


End Sub

Private Sub Command5_Click()
'上一本
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then Adodc1.Recordset.MoveFirst
xh = Adodc1.Recordset.Fields("图书名称")

End Sub

Private Sub Command6_Click()
'下一本
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast
xh = Adodc1.Recordset.Fields("图书名称")

End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""

End Sub

Private Sub 个人信息_Click()
Load Form5
Form5.Show
End Sub

Private Sub 退出_Click()
yn2 = MsgBox("确实退出登录吗(Y/N)?", vbYesNo, "退出登录提示")
If yn2 = 6 Then
    For i = 0 To Forms.count - 1
        Unload Form4
    Next i
End If

'Unload Form4
'Load Form1
'Form1.Show
End Sub

Private Sub 我_Click(Index As Integer)
MsgBox " '" & username & " '", , "提示"
End Sub

Private Sub 我的订单_Click()
Load Form6
Form6.Show
End Sub
