VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "学生信息"
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10200
   LinkTopic       =   "Form3"
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   6255
   ScaleWidth      =   10200
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command8 
      Caption         =   "上一条"
      Height          =   420
      Left            =   8745
      TabIndex        =   28
      Top             =   5670
      Width           =   1005
   End
   Begin VB.CommandButton Command7 
      Caption         =   "下一条"
      Height          =   420
      Left            =   6270
      TabIndex        =   27
      Top             =   5670
      Width           =   1005
   End
   Begin VB.CommandButton Command6 
      Caption         =   "最后一条"
      Height          =   420
      Left            =   3630
      TabIndex        =   26
      Top             =   5670
      Width           =   1005
   End
   Begin VB.CommandButton Command5 
      Caption         =   "第一条"
      Height          =   420
      Left            =   1155
      TabIndex        =   25
      Top             =   5670
      Width           =   1005
   End
   Begin VB.CommandButton Command4 
      Caption         =   "刷新"
      Height          =   420
      Left            =   9075
      TabIndex        =   24
      Top             =   4860
      Width           =   675
   End
   Begin VB.CommandButton Command3 
      Caption         =   "修改"
      Height          =   420
      Left            =   6435
      TabIndex        =   23
      Top             =   4860
      Width           =   675
   End
   Begin VB.CommandButton Command2 
      Caption         =   "删除"
      Height          =   420
      Left            =   3795
      TabIndex        =   22
      Top             =   4860
      Width           =   675
   End
   Begin VB.CommandButton Command1 
      Caption         =   "添加"
      Height          =   420
      Left            =   1155
      TabIndex        =   21
      Top             =   4860
      Width           =   675
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "G:\Game\Manage\data\test.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   555
      Left            =   6930
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "kecheng"
      Top             =   3915
      Width           =   2820
   End
   Begin VB.TextBox Text9 
      DataField       =   "xuefen"
      DataSource      =   "Data3"
      Height          =   420
      Left            =   7755
      TabIndex        =   20
      Top             =   2295
      Width           =   1830
   End
   Begin VB.TextBox Text8 
      DataField       =   "cName"
      DataSource      =   "Data3"
      Height          =   420
      Left            =   7755
      TabIndex        =   19
      Top             =   1620
      Width           =   1830
   End
   Begin VB.Frame Frame3 
      Caption         =   "课程"
      Height          =   3120
      Left            =   6930
      TabIndex        =   14
      Top             =   405
      Width           =   2820
      Begin VB.TextBox Text7 
         DataField       =   "cNo"
         DataSource      =   "Data3"
         Height          =   420
         Left            =   825
         TabIndex        =   15
         Top             =   540
         Width           =   1830
      End
      Begin VB.Label Label9 
         Caption         =   "学分"
         Height          =   420
         Left            =   165
         TabIndex        =   18
         Top             =   2025
         Width           =   675
      End
      Begin VB.Label Label8 
         Caption         =   "课名"
         Height          =   285
         Left            =   165
         TabIndex        =   17
         Top             =   1350
         Width           =   510
      End
      Begin VB.Label Label7 
         Caption         =   "课号"
         Height          =   285
         Left            =   165
         TabIndex        =   16
         Top             =   675
         Width           =   675
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "G:\Game\Manage\data\test.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   555
      Left            =   3630
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "student"
      Top             =   3915
      Width           =   2820
   End
   Begin VB.TextBox Text6 
      DataField       =   "score"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   4455
      TabIndex        =   13
      Top             =   2295
      Width           =   1830
   End
   Begin VB.TextBox Text5 
      DataField       =   "sName"
      DataSource      =   "Data2"
      Height          =   420
      Left            =   4455
      TabIndex        =   12
      Top             =   1620
      Width           =   1830
   End
   Begin VB.Frame Frame2 
      Caption         =   "学生"
      Height          =   3120
      Left            =   3630
      TabIndex        =   7
      Top             =   405
      Width           =   2820
      Begin VB.TextBox Text4 
         DataField       =   "sNo"
         DataSource      =   "Data2"
         Height          =   420
         Left            =   825
         TabIndex        =   8
         Top             =   540
         Width           =   1830
      End
      Begin VB.Label Label4 
         Caption         =   "学号"
         Height          =   285
         Left            =   165
         TabIndex        =   11
         Top             =   675
         Width           =   675
      End
      Begin VB.Label Label5 
         Caption         =   "姓名"
         Height          =   285
         Left            =   165
         TabIndex        =   10
         Top             =   1350
         Width           =   675
      End
      Begin VB.Label Label6 
         Caption         =   "年龄"
         Height          =   420
         Left            =   165
         TabIndex        =   9
         Top             =   2025
         Width           =   675
      End
   End
   Begin VB.TextBox Text3 
      DataField       =   "score"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   1155
      TabIndex        =   3
      Top             =   2295
      Width           =   1830
   End
   Begin VB.TextBox Text2 
      DataField       =   "cNo"
      DataSource      =   "Data1"
      Height          =   420
      Left            =   1155
      TabIndex        =   2
      Top             =   1620
      Width           =   1830
   End
   Begin VB.Frame Frame1 
      Caption         =   "成绩"
      Height          =   3120
      Left            =   330
      TabIndex        =   0
      Top             =   405
      Width           =   2820
      Begin VB.TextBox Text1 
         DataField       =   "sNo"
         DataSource      =   "Data1"
         Height          =   420
         Left            =   825
         TabIndex        =   1
         Top             =   540
         Width           =   1830
      End
      Begin VB.Label Label3 
         Caption         =   "成绩"
         Height          =   420
         Left            =   165
         TabIndex        =   6
         Top             =   2025
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "课号"
         Height          =   285
         Left            =   165
         TabIndex        =   5
         Top             =   1350
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "学号"
         Height          =   285
         Left            =   165
         TabIndex        =   4
         Top             =   675
         Width           =   675
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "G:\Game\Manage\data\test.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   555
      Left            =   330
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "score"
      Top             =   3915
      Width           =   2985
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "查询："
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   330
      TabIndex        =   29
      Top             =   5670
      Width           =   675
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    With Data1.Recordset
         .AddNew
         .Fields("sno") = "6"
         .Fields("cno") = "c1"
         .Fields("score") = 97
         .Update
    End With
    Form_Activate
End Sub

Private Sub Command2_Click()
    Data1.Recordset.Delete
    Form_Activate
End Sub

Private Sub Command3_Click()
    With Data1.Recordset
         .Edit
         .Fields("sno") = "6"
         .Fields("cno") = "c1"
         .Fields("score") = 100
         .Update
    End With
    Form_Activate
End Sub


Private Sub Command4_Click()
    Data1.Refresh
End Sub

Private Sub Command5_Click()
    Data1.Recordset.FindFirst "sno='2'"
End Sub

Private Sub Command6_Click()
    Data1.Recordset.FindLast "sno='2'"
End Sub

Private Sub Command7_Click()
    Data1.Recordset.FindNext "sno='2'"
End Sub

Private Sub Command8_Click()
    Data1.Recordset.FindPrevious "sno='2'"
End Sub

Private Sub Form_Activate()
    Data1.Recordset.MoveLast
    Data1.Caption = "记录数为：" & Data1.Recordset.RecordCount
    
    Data2.Recordset.MoveLast
    Data2.Caption = "记录数为：" & Data2.Recordset.RecordCount
    
    Data3.Recordset.MoveLast
    Data3.Caption = "记录数为：" & Data3.Recordset.RecordCount
End Sub

