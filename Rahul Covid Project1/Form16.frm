VERSION 5.00
Begin VB.Form Form16 
   Caption         =   "Form16"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form16"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   2655
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4815
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1920
         TabIndex        =   7
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1920
         TabIndex        =   6
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   1920
         TabIndex        =   5
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Updated Records In the System"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   11
         Top             =   120
         Width           =   2535
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total Cases"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total Recovered"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Active Cases"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Width           =   1215
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Index           =   0
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form16.frx":0000
      Top             =   5640
      Width           =   10455
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   6240
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form16.frx":028E
      Top             =   240
      Width           =   10095
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "About us"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "About us"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "India will Fight Agaist Covid 19"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   15135
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4935
      Index           =   2
      Left            =   0
      Picture         =   "Form16.frx":06EE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5055
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4935
      Index           =   1
      Left            =   0
      Picture         =   "Form16.frx":1C7ED
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5055
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4935
      Index           =   0
      Left            =   0
      Picture         =   "Form16.frx":388EC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

