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
      TabIndex        =   16
      Top             =   4920
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   735
      Left            =   2040
      TabIndex        =   15
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
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
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
         TabIndex        =   14
         Text            =   "Text6"
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   5160
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   1920
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   1320
         TabIndex        =   12
         Text            =   "Text4"
         Top             =   2880
         Width           =   2415
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   1320
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1320
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1320
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   360
         Width           =   6255
      End
      Begin VB.ComboBox combo1 
         Height          =   300
         Left            =   5160
         TabIndex        =   8
         Text            =   "选择类别"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "出版时间"
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
         Caption         =   "登记时间"
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

End Sub

Private Sub Form_Load()

Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""

End Sub
