VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   Caption         =   "管理系统"
   ClientHeight    =   6765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11745
   BeginProperty Font 
      Name            =   "幼圆"
      Size            =   15
      Charset         =   134
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   6765
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      IMEMode         =   3  'DISABLE
      Left            =   7095
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2835
      Width           =   3810
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5010
      Left            =   6105
      TabIndex        =   0
      Top             =   945
      Width           =   5130
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C000C0&
         Caption         =   "登录"
         Height          =   555
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3240
         Width           =   4635
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         IMEMode         =   3  'DISABLE
         Left            =   990
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   945
         Width           =   3810
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "没有账号？立即注册"
         BeginProperty Font 
            Name            =   "幼圆"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   420
         Left            =   165
         TabIndex        =   7
         Top             =   4455
         Width           =   4635
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "密码"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   420
         Left            =   165
         TabIndex        =   3
         Top             =   2025
         Width           =   675
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "账号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   285
         Left            =   165
         TabIndex        =   2
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "欢迎登录"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   555
         Left            =   1815
         TabIndex        =   1
         Top             =   135
         Width           =   1500
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "刘鑫磊@版权所有"
      BeginProperty Font 
         Name            =   "幼圆"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3465
      TabIndex        =   8
      Top             =   6345
      Width           =   3645
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If Text1.Text = 123 And Text2.Text = 123 Then
        Dim a1 As String
        a1 = MsgBox("恭喜您，登录成功！", vbExclamation, "提示")
        Form1.Visible = False
        Form3.Visible = True
    ElseIf Text1.Text = 456 And Text2.Text = 456 Then
        Dim a2 As String
        a2 = MsgBox("恭喜您，登录成功！", vbExclamation, "提示")
        Form1.Visible = False
        Form3.Visible = True
    ElseIf Text1.Text = 789 And Text2.Text = 789 Then
        Dim a3 As String
        a3 = MsgBox("恭喜您，登录成功！", vbExclamation, "提示")
        Form1.Visible = False
        Form3.Visible = True
    Else
        MsgBox ("密码错误！")
    End If
    
End Sub

Private Sub Label4_Click()
    Form2.Visible = True
    Form1.Visible = False
End Sub

