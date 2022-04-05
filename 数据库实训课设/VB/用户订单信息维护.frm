VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form6 
   Caption         =   "用户订单"
   ClientHeight    =   7845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14910
   LinkTopic       =   "Form6"
   ScaleHeight     =   7845
   ScaleWidth      =   14910
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "下一本"
      Height          =   735
      Left            =   3240
      TabIndex        =   6
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "上一本"
      Height          =   735
      Left            =   480
      TabIndex        =   5
      Top             =   6600
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "退出"
      Height          =   735
      Left            =   12600
      TabIndex        =   4
      Top             =   6600
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   5655
      Left            =   480
      ScaleHeight     =   5595
      ScaleWidth      =   13755
      TabIndex        =   2
      Top             =   480
      Width           =   13815
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   495
         Left            =   480
         Top             =   240
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
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "用户订单信息维护.frx":0000
         Height          =   3975
         Left            =   480
         TabIndex        =   3
         Top             =   960
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   7011
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
         Caption         =   "订单浏览"
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
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消订单"
      Height          =   735
      Left            =   9240
      TabIndex        =   1
      Top             =   6600
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "计算总价"
      Height          =   735
      Left            =   6000
      TabIndex        =   0
      Top             =   6600
      Width           =   2655
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Dim cn As New ADODB.Connection
Dim rs As New ADODB.Connection
'Dim sql As String
Dim sum As Integer
cn.Open "provider=SQLOLEDB;Data source=DESKTOP-KLVGTUE;User Id=ydczsq; password='881212';Inital Catalog=bookstore"
Do While Not Adodc1.Recordset.EOF
    sum = sum + Adodc1.Recordset("价格")
    Adodc1.Recordset.MoveNext
Loop
MsgBox " 总计为" & sum & "元 ", , "总价"


'Adodc1.ConnectionString = "dsn=1"
'Adodc1.CommandType = adCmdUnknown
'Adodc1.RecordSource = "select * from shopping where "
End Sub

Private Sub Command2_Click()
x = MsgBox("确定要取消该订单吗？", vbYesNo + vbQuestion, "提示")
If x = vbYes Then
    Adodc1.Recordset.Delete
    'Adodc1.Recordset.MoveNext
    Adodc1.Refresh
End If
End Sub

Private Sub Command3_Click()
Unload Form6
Load Form4
Form4.Show
End Sub

Private Sub Command4_Click()

Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then Adodc1.Recordset.MoveFirst


End Sub

Private Sub Command5_Click()

Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast
End Sub

Private Sub Form_Load()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Connection
Dim sql As String
cn.Open "provider=SQLOLEDB;Data source=DESKTOP-KLVGTUE;User Id=ydczsq; password='881212';Inital Catalog=bookstore"
Adodc1.ConnectionString = "dsn=1"
Adodc1.CommandType = adCmdUnknown
Adodc1.RecordSource = "select * from shopping where 用户姓名='" & username & " '"
Adodc1.Refresh
End Sub
