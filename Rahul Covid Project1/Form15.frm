VERSION 5.00
Begin VB.Form Form15 
   Caption         =   "Form15"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form15"
   Picture         =   "Form15.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   1320
      TabIndex        =   4
      Top             =   7560
      Width           =   12975
      Begin VB.CommandButton Command3 
         Caption         =   "Forgot  Mobile Number"
         Height          =   615
         Left            =   9480
         TabIndex        =   7
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Submit"
         Height          =   555
         Left            =   5640
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Back"
         Height          =   615
         Left            =   360
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   2520
      TabIndex        =   1
      Top             =   2160
      Width           =   10575
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   5520
         TabIndex        =   3
         Top             =   1560
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Enter The Mobile Number"
         Height          =   735
         Left            =   1080
         TabIndex        =   2
         Top             =   1560
         Width           =   3135
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter the Required Details , So that We will help you to Recover the Username"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2520
      TabIndex        =   0
      Top             =   600
      Width           =   11055
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Form15.Hide
Form1.Show
End Sub

Private Sub Command2_Click()
If Text1.Text = "9860552959" Then
MsgBox ("Your Username is = admin")

End If
If Text1.Text = "9139211819" Then
MsgBox ("Your Username is= rahul9")
End If
If Text1.Text = "9021003529" Then
MsgBox ("Your Username is = raj20")
End If

End Sub

Private Sub Command3_Click()
MsgBox ("Please Contact to It Team to recover your Credentials")

End Sub

Private Sub Form_Load()
Form15.Picture = LoadPicture("C:\Rahul Covid Project1\Resources\ee.jpg")
End Sub
