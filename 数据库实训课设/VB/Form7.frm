VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form7 
   Caption         =   "管理员界面"
   ClientHeight    =   12000
   ClientLeft      =   225
   ClientTop       =   915
   ClientWidth     =   19710
   LinkTopic       =   "Form7"
   ScaleHeight     =   12000
   ScaleWidth      =   19710
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Height          =   10695
      Index           =   1
      Left            =   720
      ScaleHeight     =   10635
      ScaleWidth      =   18195
      TabIndex        =   9
      Top             =   720
      Width           =   18255
      Begin VB.PictureBox Picture2 
         Height          =   1335
         Left            =   2160
         ScaleHeight     =   1275
         ScaleWidth      =   13275
         TabIndex        =   18
         Top             =   9000
         Width           =   13335
         Begin VB.CommandButton Command6 
            Caption         =   "取消"
            Height          =   615
            Left            =   10920
            TabIndex        =   22
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton Command5 
            Caption         =   "删除图书"
            Height          =   615
            Left            =   8040
            TabIndex        =   21
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton Command4 
            Caption         =   "添加图书"
            Height          =   615
            Left            =   4680
            TabIndex        =   20
            Top             =   360
            Width           =   1815
         End
         Begin VB.CommandButton Command3 
            Caption         =   "修改图书"
            Height          =   615
            Left            =   1320
            TabIndex        =   19
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "按图书信息查询"
         Height          =   1215
         Index           =   1
         Left            =   2160
         TabIndex        =   12
         Top             =   1200
         Width           =   13335
         Begin VB.TextBox Text4 
            DataField       =   "图书名称"
            DataSource      =   "Adodc1"
            Height          =   495
            Index           =   1
            Left            =   1920
            TabIndex        =   15
            Text            =   "Text4"
            Top             =   480
            Width           =   2055
         End
         Begin VB.TextBox Text5 
            DataField       =   "作者"
            DataSource      =   "Adodc1"
            Height          =   495
            Index           =   1
            Left            =   6960
            TabIndex        =   14
            Text            =   "Text5"
            Top             =   480
            Width           =   2295
         End
         Begin VB.CommandButton Command2 
            Caption         =   "查询"
            Height          =   615
            Index           =   1
            Left            =   10800
            TabIndex        =   13
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "书  名："
            Height          =   495
            Index           =   1
            Left            =   720
            TabIndex        =   17
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "作  者："
            Height          =   375
            Index           =   1
            Left            =   5880
            TabIndex        =   16
            Top             =   480
            Width           =   1095
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form7.frx":0000
         Height          =   3615
         Index           =   1
         Left            =   2640
         TabIndex        =   10
         Top             =   4560
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   6376
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
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
      Begin VB.Frame Frame3 
         Caption         =   "图书信息"
         Height          =   4695
         Index           =   1
         Left            =   2160
         TabIndex        =   11
         Top             =   3960
         Width           =   13335
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   495
            Left            =   10920
            Top             =   120
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   873
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
            RecordSource    =   "books"
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
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   10695
      Index           =   0
      Left            =   720
      ScaleHeight     =   10635
      ScaleWidth      =   18195
      TabIndex        =   0
      Top             =   1320
      Width           =   18255
      Begin VB.PictureBox Picture3 
         Height          =   975
         Left            =   2520
         ScaleHeight     =   915
         ScaleWidth      =   13275
         TabIndex        =   23
         Top             =   8280
         Width           =   13335
         Begin VB.CommandButton Command8 
            Caption         =   "取消"
            Height          =   615
            Left            =   9240
            TabIndex        =   25
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton Command7 
            Caption         =   "修改订单"
            Height          =   615
            Left            =   3720
            TabIndex        =   24
            Top             =   240
            Width           =   1815
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form7.frx":0015
         Height          =   3615
         Index           =   0
         Left            =   3120
         TabIndex        =   1
         Top             =   3360
         Width           =   12375
         _ExtentX        =   21828
         _ExtentY        =   6376
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
      Begin VB.Frame Frame3 
         Caption         =   "订单信息"
         Height          =   4695
         Index           =   0
         Left            =   2520
         TabIndex        =   8
         Top             =   2640
         Width           =   13335
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   495
            Left            =   10920
            Top             =   120
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   873
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
      End
      Begin VB.Frame Frame1 
         Caption         =   "按用户信息查询"
         Height          =   1215
         Index           =   0
         Left            =   2520
         TabIndex        =   2
         Top             =   600
         Width           =   13335
         Begin VB.CommandButton Command1 
            Caption         =   "查询"
            Height          =   615
            Index           =   0
            Left            =   10800
            TabIndex        =   7
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox Text2 
            DataSource      =   "Adodc2"
            Height          =   495
            Index           =   0
            Left            =   6960
            TabIndex        =   6
            Text            =   "Text2"
            Top             =   480
            Width           =   2295
         End
         Begin VB.TextBox Text1 
            Height          =   495
            Index           =   0
            Left            =   1920
            TabIndex        =   4
            Text            =   "Text1"
            Top             =   480
            Width           =   2055
         End
         Begin VB.Label Label2 
            Caption         =   "姓  名："
            Height          =   495
            Index           =   0
            Left            =   720
            TabIndex        =   5
            Top             =   480
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "电  话："
            Height          =   375
            Index           =   0
            Left            =   5880
            TabIndex        =   3
            Top             =   480
            Width           =   855
         End
      End
   End
   Begin VB.Menu 管理员 
      Caption         =   "管理员"
   End
   Begin VB.Menu 订单管理 
      Caption         =   "订单管理"
      Begin VB.Menu 查询订单 
         Caption         =   "查询订单"
         Index           =   0
      End
      Begin VB.Menu 修改订单 
         Caption         =   "修改订单"
         Index           =   0
      End
      Begin VB.Menu 统计销量 
         Caption         =   "统计销量"
         Index           =   0
      End
   End
   Begin VB.Menu 图书管理 
      Caption         =   "图书管理"
      Begin VB.Menu 查询图书 
         Caption         =   "查询图书"
         Index           =   0
      End
      Begin VB.Menu 添加图书 
         Caption         =   "添加图书"
         Index           =   0
      End
      Begin VB.Menu 修改图书 
         Caption         =   "修改图书"
         Index           =   0
      End
      Begin VB.Menu 删除图书 
         Caption         =   "删除图书"
         Index           =   0
      End
      Begin VB.Menu 统计图书数量 
         Caption         =   "统计图书数量"
         Index           =   0
      End
   End
   Begin VB.Menu 退出 
      Caption         =   "退出"
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub 查询图书信息_Click(Index As Integer)

Picture1(1).Visible = True
Picture1(0).Visible = False

End Sub

Private Sub Command1_Click(Index As Integer)
'按照用户信息查询订单信息 姓名和电话
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Connection
cn.Open "provider=SQLOLEDB;Data source=LAPTOP-KEK71Q0C;User Id=ydczsq; password='991212';Inital Catalog=bookstore"
If Text1(0).Text <> "" And Text2(0).Text <> "" Then
    cn.Execute "insert into AA select  from shopping where 姓名=' " & Trim(Text1(0).Text) & " ' and 电话=' " & Trim(Text2(0).Text) & " '", cn
ElseIf Text1(0).Text <> "" Then
    cn.Execute "insert into AA select * from shopping where 姓名=' " & Trim(Text1(0).Text) & " '", cn
ElseIf Text2(0).Text <> "" Then
    cn.Execute "insert into AA select * from shopping where 电话=' " & Trim(Text2(0).Text) & " '", cn
Else
    MsgBox "无法查询，请重新输入所需信息！", , "提示"
End If
Adodc2.Refresh
Form7.Refresh

End Sub

Private Sub Command10_Click()
'按图书信息查询图书信息 书名和作者
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Connection
cn.Open "provider=SQLOLEDB;Data source=DESKTOP-KLVGTUE;User Id=ydczsq; password='881212';Inital Catalog=bookstore"
If Text4(1).Text <> "" And Text5(1).Text <> "" Then
    cn.Execute "insert into BB select * from books where 书名=' " & Trim(Text4(1).Text) & " ' and 作者=' " & Trim(Text5(1).Text) & " '", cn
ElseIf Text4(1).Text <> "" Then
    cn.Execute "insert into BB select * from books where 书名=' " & Trim(Text4(1).Text) & " '", cn
ElseIf Text5(1).Text <> "" Then
    cn.Execute "insert into BB select * from books where 作者=' " & Trim(Text5(1).Text) & " '", cn
Else
    MsgBox "无法查询，请重新输入所需信息！", , "提示"
End If
Adodc1.Refresh
Form7.Refresh
End Sub


Private Sub Command2_Click(Index As Integer)

'按图书信息查询图书信息 书名和作者
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Connection
cn.Open "provider=SQLOLEDB;Data source=DESKTOP-KLVGTUE;User Id=ydczsq; password='881212';Inital Catalog=bookstore"
If Text4(1).Text <> "" And Text5(1).Text <> "" Then
    cn.Execute "insert into AA select * from books where 图书名称=' " & Trim(Text4(1).Text) & " ' and 作者=' " & Trim(Text5(1).Text) & " '", cn
ElseIf Text4(1).Text <> "" Then
    cn.Execute "insert into AA select * from orders where 图书名称=' " & Trim(Text4(1).Text) & " '", cn
ElseIf Text5(1).Text <> "" Then
    cn.Execute "insert into AA select * from orders where 图书名称=' " & Trim(Text5(1).Text) & " '", cn
Else
    MsgBox "无法查询，请重新输入所需信息！", , "提示"
End If
Adodc2.Refresh
Form7.Refresh

End Sub

Private Sub Command4_Click()
'添加图书
Form8.Show
Adodc1.Refresh
Form7.Refresh

End Sub

Private Sub Command5_Click()
'删除图书
Adodc1.Recordset.Delete
'Adodc1.Recordset.MoveNext
Adodc1.Refresh

End Sub

Private Sub Command6_Click()
sign_out = MsgBox("是否退出系统(Y/N)？", vbYesNo, "提示")
If sign_out = 6 Then
    For i = 0 To Forms.count - 1
        Unload Forms(0)
    Next i
End If

End Sub

Private Sub Form_Load()
Picture1(1).Visible = True
Picture1(0).Visible = False

End Sub

Private Sub Frame3_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

End Sub

Private Sub 查询订单_Click(Index As Integer)
Picture1(0).Visible = True
Picture1(1).Visible = False
End Sub

Private Sub 查询图书_Click(Index As Integer)
Picture1(1).Visible = True
Picture1(0).Visible = False

End Sub

Private Sub 删除图书_Click(Index As Integer)
Adodc1.Recordset.Delete
'Adodc1.Recordset.MoveNext
Adodc1.Refresh
End Sub

Private Sub 添加图书_Click(Index As Integer)
Form8.Show
Adodc1.Refresh
Form7.Refresh
End Sub

Private Sub 统计图书数量_Click(Index As Integer)
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Connection
Dim count As Integer
count = 0
cn.Open "provider=SQLOLEDB;Data source=LAPTOP-KEK71Q0C;User Id=ydczsq; password='991212';Inital Catalog=bookstore"
rs.Open "select * from Books", cn
Do While Not rs.EOF
    count = count + 1
Loop
MsgBox "当前存有图书 ' &count& '本", , "图书总数"



End Sub

Private Sub 统计销量_Click(Index As Integer)
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Connection
Dim count As Integer
count = 0
cn.Open "provider=SQLOLEDB;Data source=LAPTOP-KEK71Q0C;User Id=ydczsq; password='991212';Inital Catalog=bookstore"
rs.Open "select * from shopping", cn
Do While Not rs.EOF
    count = count + 1
Loop
MsgBox "当前存有图书 ' &count& '本", , "图书总数"


End Sub

Private Sub 退出_Click()

sign_out = MsgBox("是否退出系统(Y/N)？", vbYesNo, "提示")
If sign_out = 6 Then
    For i = 0 To Forms.count - 1
        Unload Forms(0)
    Next i
End If


End Sub

