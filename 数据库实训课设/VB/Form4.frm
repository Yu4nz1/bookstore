VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "�û�����"
   ClientHeight    =   9060
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   17295
   LinkTopic       =   "Form4"
   ScaleHeight     =   9060
   ScaleWidth      =   17295
   StartUpPosition =   3  '����ȱʡ
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
      Connect         =   "DSN=�鱾��Ϣ"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "�鱾��Ϣ"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "shopping"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
      Caption         =   "��һ��"
      Height          =   735
      Left            =   12240
      TabIndex        =   15
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "��һ��"
      Height          =   735
      Left            =   9840
      TabIndex        =   14
      Top             =   7920
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "���붩��"
      Height          =   735
      Left            =   14400
      TabIndex        =   1
      Top             =   7920
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   6855
      Left            =   720
      ScaleHeight     =   6795
      ScaleWidth      =   15795
      TabIndex        =   0
      Top             =   840
      Width           =   15855
      Begin VB.TextBox Text5 
         DataField       =   "���"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1680
         TabIndex        =   16
         Top             =   1320
         Width           =   2415
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   840
         Top             =   2160
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
         Connect         =   "DSN=�鱾��Ϣ"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "�鱾��Ϣ"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Books"
         Caption         =   "Adodc1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
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
         Caption         =   "ȡ��"
         Height          =   735
         Left            =   12360
         TabIndex        =   13
         Top             =   2040
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         Caption         =   "ȷ��"
         Height          =   735
         Left            =   8280
         TabIndex        =   12
         Top             =   2040
         Width           =   2175
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Form4.frx":0000
         Height          =   3255
         Left            =   720
         TabIndex        =   11
         Top             =   3240
         Width           =   14655
         _ExtentX        =   25850
         _ExtentY        =   5741
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   13
         FormatLocked    =   -1  'True
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   "ISBN��"
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
            Caption         =   "����"
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
         BeginProperty Column02 
            DataField       =   ""
            Caption         =   "����"
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
         BeginProperty Column03 
            DataField       =   ""
            Caption         =   "������"
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
         BeginProperty Column04 
            DataField       =   ""
            Caption         =   "�۸�"
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
         BeginProperty Column05 
            DataField       =   ""
            Caption         =   "����ʱ��"
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
               ColumnWidth     =   2445.166
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2294.929
            EndProperty
            BeginProperty Column02 
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3105.071
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1679.811
            EndProperty
            BeginProperty Column05 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox Text4 
         DataField       =   "������"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   10200
         TabIndex        =   10
         Top             =   1320
         Width           =   4815
      End
      Begin VB.TextBox Text3 
         DataField       =   "����"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   5760
         TabIndex        =   9
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         DataField       =   "ISBN��"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   10200
         TabIndex        =   8
         Top             =   480
         Width           =   4815
      End
      Begin VB.TextBox Text1 
         DataField       =   "ͼ������"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   1680
         TabIndex        =   7
         Top             =   480
         Width           =   6375
      End
      Begin VB.CheckBox Check5 
         Caption         =   "������"
         Height          =   420
         Left            =   9000
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox Check4 
         Caption         =   "ISBN��"
         Height          =   495
         Left            =   9000
         TabIndex        =   5
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox Check3 
         Caption         =   "��  ��"
         Height          =   420
         Left            =   4560
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         Caption         =   "��  ��"
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "��  ��"
         Height          =   420
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Menu �� 
      Caption         =   "��"
      Index           =   0
      Visible         =   0   'False
   End
   Begin VB.Menu ������Ϣ 
      Caption         =   "������Ϣ"
   End
   Begin VB.Menu �ҵĶ��� 
      Caption         =   "�ҵĶ���"
   End
   Begin VB.Menu �˳� 
      Caption         =   "�˳�"
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
yn2 = MsgBox("ȷʵ�˳�ϵͳ��(Y/N)?", vbYesNo, "�˳�ϵͳ��ʾ")
If yn2 = 6 Then
For i = 0 To Forms.count - 1
  Unload Forms(0)
Next i
End If
End Sub

Private Sub Command3_Click()
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
cn.Open "provider=SQLOLEDB;Data source=LAPTOP-KEK71Q0C;User Id=ydczsq; password='991212';Inital Catalog=Bookstore"
cn.Execute "insert into shopping(ͼ������,ISBN��,����,������,���) values('" & Text1.Text & "','" & Text2.Text & "','" & Text3.Text & "','" & Text4.Text & "','" & Text5.Text & "')"
Form1.Refresh
Adodc2.Refresh

End Sub

Private Sub Command5_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Adodc1.Recordset.MovePrevious
If Adodc1.Recordset.BOF Then Adodc1.Recordset.MoveFirst
xh = Adodc1.Recordset.Fields("ͼ������")

End Sub

Private Sub Command6_Click()
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast
xh = Adodc1.Recordset.Fields("ͼ������")

End Sub

Private Sub Form_Load()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""

End Sub

Private Sub ������Ϣ_Click()
Load Form5
Form5.Show
End Sub

Private Sub �˳�_Click()
Unload Form4
Load Form1
Form1.Show
End Sub

Private Sub �ҵĶ���_Click()
Load Form6
Form6.Show
End Sub
