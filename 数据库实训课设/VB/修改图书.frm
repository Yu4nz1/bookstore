VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form Form2 
   Caption         =   "修改图书"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10305
   LinkTopic       =   "Form2"
   ScaleHeight     =   6480
   ScaleWidth      =   10305
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command4 
      Caption         =   "退出"
      Height          =   735
      Left            =   8280
      TabIndex        =   18
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "修改"
      Height          =   735
      Left            =   5760
      TabIndex        =   17
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "下一本"
      Height          =   735
      Left            =   3120
      TabIndex        =   16
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "上一本"
      Height          =   735
      Left            =   720
      TabIndex        =   15
      Top             =   5400
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      Height          =   4695
      Left            =   840
      ScaleHeight     =   4635
      ScaleWidth      =   8595
      TabIndex        =   0
      Top             =   360
      Width           =   8655
      Begin VB.TextBox Text1 
         DataField       =   "图书名称"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   6255
      End
      Begin VB.TextBox Text2 
         DataField       =   "作者"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1320
         TabIndex        =   6
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         DataField       =   "ISBN号"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1320
         TabIndex        =   5
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         DataField       =   "出版社"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1320
         TabIndex        =   4
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         DataField       =   "价格"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   5160
         TabIndex        =   3
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         DataField       =   "出版时间"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   5160
         TabIndex        =   2
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox Text7 
         DataField       =   "类别"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   5160
         TabIndex        =   1
         Top             =   1200
         Width           =   2415
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   450
         Left            =   6120
         Top             =   3600
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
      Begin VB.Label Label1 
         Caption         =   "书  名"
         Height          =   375
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "作  者"
         Height          =   375
         Left            =   360
         TabIndex        =   13
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "ISBN号"
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "出版社"
         Height          =   495
         Left            =   360
         TabIndex        =   11
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "出版时间"
         Height          =   495
         Left            =   4200
         TabIndex        =   10
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "类  别"
         Height          =   375
         Left            =   4200
         TabIndex        =   9
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "单  价"
         Height          =   495
         Left            =   4200
         TabIndex        =   8
         Top             =   2040
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'上一本
Text1.Enabled = True
Text2.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then Adodc1.Recordset.MoveFirst
xh = Adodc1.Recordset.Fields("图书名称")

End Sub

Private Sub Command2_Click()
'下一本
Text1.Enabled = True
Text2.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast
xh = Adodc1.Recordset.Fields("图书名称")

End Sub

Private Sub Command3_Click()
Text1.Enabled = True
Text2.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True

Adodc1.Recordset.Update


End Sub

Private Sub Command4_Click()
Unload Form2
Form7.Show
End Sub

Private Sub Form_Load()

Text1.Enabled = True
Text2.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
End Sub
