VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11340
   LinkTopic       =   "Form5"
   Picture         =   "Form5.frx":0000
   ScaleHeight     =   7830
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command3 
      Caption         =   "Back"
      Height          =   855
      Left            =   5160
      TabIndex        =   3
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Click here If you have forgot your Password"
      Height          =   975
      Left            =   6600
      TabIndex        =   2
      Top             =   2520
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click here If you have forgot your Username"
      Height          =   975
      Left            =   2400
      TabIndex        =   1
      Top             =   2520
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Be rest Assured . We are always hear to help you"
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
      Left            =   3360
      TabIndex        =   0
      Top             =   480
      Width           =   7335
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
Form6.Show



End Sub

Private Sub Command2_Click()
Unload Me
Form15.Show
End Sub

Private Sub Command3_Click()
Form1.Show
End Sub

Private Sub Form_Load()
Form5.Picture = LoadPicture("C:\Users\Admin\Desktop\BACKGROUND\aHLIdT.jpg")

End Sub
