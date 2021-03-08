VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   5700
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9060
   LinkTopic       =   "Form6"
   Picture         =   "Form6.frx":0000
   ScaleHeight     =   5700
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   735
      Left            =   8520
      TabIndex        =   8
      Top             =   7560
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   7680
      TabIndex        =   5
      Top             =   5520
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   7800
      TabIndex        =   4
      Top             =   3240
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   7800
      TabIndex        =   3
      Top             =   1560
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      Height          =   615
      Left            =   3960
      TabIndex        =   2
      Top             =   7680
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Please Enter The Required Credentials to know your Password"
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
      Left            =   3840
      TabIndex        =   7
      Top             =   120
      Width           =   7815
   End
   Begin VB.Label Label2 
      Caption         =   "Confirm the Mobile Number "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   6
      Top             =   5280
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enter the Mobile Number to know the Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   3240
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enter the Username that you want to know the Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Index           =   9
      Left            =   1800
      TabIndex        =   0
      Top             =   1440
      Width           =   4575
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "admin" And Text2.Text = "9860552959" And Text3.Text = "9860552959" Then
MsgBox ("Your Password is =admin")
End If
If Text1.Text = "rahul9" And Text2.Text = "9139211819" And Text3.Text = "9139211819" Then
MsgBox ("Your Password is =rahul9")
End If

If Text1.Text = "raj20" And Text2.Text = "9860552959" And Text3.Text = "9860552959" Then
MsgBox ("Your Password is=raj20")
End If
End Sub

Private Sub Command2_Click()
Unload Me
Form1.Show

End Sub

Private Sub Form_Load()
Form6.Picture = LoadPicture("C:\Rahul Covid Project1\Resources\geGD84.jpg")
End Sub
