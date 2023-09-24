VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "账号注册"
   ClientHeight    =   5415
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9465
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   5415
   ScaleWidth      =   9465
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C000C0&
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   5775
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3645
      Width           =   1500
   End
   Begin VB.TextBox Text2 
      Height          =   555
      IMEMode         =   3  'DISABLE
      Left            =   3465
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   2160
      Width           =   3810
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5010
      Left            =   2475
      TabIndex        =   0
      Top             =   270
      Width           =   5130
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C000C0&
         Caption         =   "确认"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   165
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   3375
         Width           =   1500
      End
      Begin VB.TextBox Text1 
         Height          =   555
         IMEMode         =   3  'DISABLE
         Left            =   990
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   945
         Width           =   3810
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H008080FF&
         Caption         =   "注册账号"
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
         TabIndex        =   4
         Top             =   135
         Width           =   1500
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
         TabIndex        =   3
         Top             =   1080
         Width           =   675
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
         TabIndex        =   2
         Top             =   2025
         Width           =   675
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    If Text1.Text = 123 And Text2.Text = 123 Then
        Dim b1 As String
        b1 = MsgBox("恭喜您，注册成功！", vbExclamation, "提示")
    ElseIf Text1.Text = 456 And Text2.Text = 456 Then
        Dim b2 As String
        b2 = MsgBox("恭喜您，注册成功！", vbExclamation, "提示")
    ElseIf Text1.Text = 789 And Text2.Text = 789 Then
        Dim b3 As String
        b3 = MsgBox("恭喜您，注册成功！", vbExclamation, "提示")
    End If
    Form2.Visible = False
    Form1.Visible = True
End Sub

Private Sub Command2_Click()
    End
End Sub
