VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H0080C0FF&
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9075
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   Picture         =   "form1.frx":0000
   ScaleHeight     =   6495
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdForgot 
      Cancel          =   -1  'True
      Caption         =   "Forgot Password?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   3840
      TabIndex        =   8
      Top             =   4920
      Width           =   1620
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8400
      Top             =   5040
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   2760
      TabIndex        =   7
      Top             =   3720
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   6360
      TabIndex        =   4
      Top             =   4920
      Width           =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   1560
      TabIndex        =   3
      Top             =   4920
      Width           =   1500
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   4320
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2520
      Width           =   3285
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   1680
      Width           =   3285
   End
   Begin VB.Label lblLabels 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Covid Center Management System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   2
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   9360
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   0
      Left            =   1560
      TabIndex        =   5
      Top             =   1680
      Width           =   1440
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Index           =   1
      Left            =   1560
      TabIndex        =   1
      Top             =   2520
      Width           =   1320
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim i As Boolean



Private Sub cmdCancel_Click(Index As Integer)
Form1.Hide
End Sub

Private Sub cmdForgot_Click(Index As Integer)
Form1.Hide
Unload Me
Form5.Show


End Sub

Private Sub cmdOK_Click(Index As Integer)
If (txtUserName = "admin" And txtPassword = "admin") Or (txtUserName = "rahul" And txtPassword = "rahul9") Or (txtUserName = "raj" And txtPassword = "raj20") Then
i = True
        ProgressBar1.Visible = True
        Else
        MsgBox "Invalid Password, try again!", , "Login"
        txtPassword.SetFocus
            
    End If
        End Sub
        
        Private Sub Timer1_Timer()
If ProgressBar1.Value <= 100 Then

        ProgressBar1.Value = ProgressBar1.Value + 5
        End If
        If ProgressBar1.Value = 100 And i Then
        Form1.Hide
        Form2.Show
        MsgBox ("Welcome Back" & txtUserName)
        Timer1.Enabled = False
        End If
End Sub

Private Sub Form_Load()
Form1.Picture = LoadPicture("C:\Users\Admin\Desktop\Girl.jpeg")
ProgressBar1.Visible = False
End Sub

