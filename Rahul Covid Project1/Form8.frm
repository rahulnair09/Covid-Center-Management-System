VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15105
   LinkTopic       =   "Form8"
   Picture         =   "Form8.frx":0000
   ScaleHeight     =   8325
   ScaleWidth      =   15105
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command15 
      Caption         =   "Add"
      Height          =   375
      Index           =   0
      Left            =   0
      TabIndex        =   64
      Top             =   7680
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   5895
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   1200
      Width           =   16095
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   2880
         TabIndex        =   50
         Top             =   3000
         Width           =   2655
         Begin VB.CheckBox Check4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Throat Pain"
            Height          =   435
            Left            =   1440
            TabIndex        =   54
            Top             =   600
            Width           =   855
         End
         Begin VB.CheckBox Check3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Tiredness"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   600
            Width           =   1095
         End
         Begin VB.CheckBox Check2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Fever"
            Height          =   255
            Left            =   1320
            TabIndex        =   52
            Top             =   240
            Width           =   1335
         End
         Begin VB.CheckBox Check1 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Coughing"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            Top             =   240
            Width           =   1455
         End
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   13680
         TabIndex        =   49
         Top             =   3600
         Width           =   1455
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   13680
         TabIndex        =   48
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   13680
         TabIndex        =   47
         Top             =   1440
         Width           =   1455
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   13680
         TabIndex        =   46
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   8520
         TabIndex        =   45
         Top             =   4800
         Width           =   2175
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   8280
         TabIndex        =   44
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   8280
         TabIndex        =   43
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   3000
         TabIndex        =   42
         Top             =   4200
         Width           =   2415
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2880
         TabIndex        =   41
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2880
         TabIndex        =   40
         Top             =   2040
         Width           =   2535
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2880
         TabIndex        =   39
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2880
         TabIndex        =   38
         Top             =   840
         Width           =   2535
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   0
         Left            =   2880
         TabIndex        =   15
         Top             =   4920
         Width           =   2655
         Begin VB.OptionButton Option2 
            Caption         =   "No"
            Height          =   315
            Left            =   1320
            TabIndex        =   56
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Yes"
            Height          =   195
            Left            =   120
            TabIndex        =   55
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFFFF&
         Height          =   735
         Index           =   0
         Left            =   8280
         TabIndex        =   14
         Top             =   480
         Width           =   2655
         Begin VB.OptionButton Option4 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pos(Mi)"
            Height          =   255
            Left            =   1440
            TabIndex        =   58
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Pos(Mj)"
            Height          =   195
            Left            =   120
            TabIndex        =   57
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   8520
         TabIndex        =   13
         Top             =   3600
         Width           =   2055
         Begin VB.OptionButton Option6 
            BackColor       =   &H00FFFFFF&
            Caption         =   "No"
            Height          =   195
            Left            =   1200
            TabIndex        =   60
            Top             =   240
            Width           =   615
         End
         Begin VB.OptionButton Option5 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Yes"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.Frame Frame7 
         Height          =   1215
         Left            =   13560
         TabIndex        =   12
         Top             =   4680
         Width           =   2295
         Begin VB.OptionButton Option9 
            Caption         =   "ICU"
            Height          =   255
            Left            =   720
            TabIndex        =   63
            Top             =   720
            Width           =   855
         End
         Begin VB.OptionButton Option8 
            Caption         =   "Special"
            Height          =   195
            Left            =   1440
            TabIndex        =   62
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option7 
            BackColor       =   &H00FFFFFF&
            Caption         =   "General"
            Height          =   375
            Left            =   240
            TabIndex        =   61
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   375
         Index           =   2
         Left            =   7560
         TabIndex        =   11
         Top             =   5880
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Symptoms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   15
         Left            =   120
         TabIndex        =   32
         Top             =   3240
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Mobile No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   14
         Left            =   120
         TabIndex        =   31
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   30
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   29
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Admission ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   27
         Top             =   4200
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Result of the Patient"
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
         Index           =   3
         Left            =   5880
         TabIndex        =   26
         Top             =   600
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Whether the Patient has Tested Positive from our Covid Center"
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
         Index           =   7
         Left            =   120
         TabIndex        =   25
         Top             =   4800
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of the Report Generated by Pathology Lab"
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
         Left            =   5760
         TabIndex        =   24
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Hospital Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   5760
         TabIndex        =   23
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Admission"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   4
         Left            =   5760
         TabIndex        =   22
         Top             =   4800
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Whether the Patient was adviced for Home Quaratine"
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
         Index           =   5
         Left            =   5760
         TabIndex        =   21
         Top             =   3600
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Respiratory Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   8
         Left            =   11280
         TabIndex        =   20
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Current Blood Pressure Rate"
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
         Index           =   10
         Left            =   11160
         TabIndex        =   19
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Sp 02 Rate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   12
         Left            =   11280
         TabIndex        =   18
         Top             =   2640
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Bed No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   11160
         TabIndex        =   17
         Top             =   3600
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ward"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   16
         Left            =   11160
         TabIndex        =   16
         Top             =   4800
         Width           =   2415
      End
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   8280
      TabIndex        =   9
      Top             =   8760
      Width           =   1935
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   0
      TabIndex        =   8
      Top             =   9480
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   4683
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Menus"
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   7320
      Width           =   15855
      Begin VB.CommandButton command7 
         Caption         =   "MoveNext"
         Height          =   375
         Index           =   0
         Left            =   9720
         TabIndex        =   36
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "MoveLast"
         Height          =   375
         Index           =   1
         Left            =   12360
         TabIndex        =   35
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "MovePrevious"
         Height          =   375
         Index           =   0
         Left            =   11040
         TabIndex        =   34
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "MoveFirst"
         Height          =   375
         Index           =   0
         Left            =   8400
         TabIndex        =   33
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command20 
         Caption         =   "View Database"
         Height          =   375
         Index           =   6
         Left            =   13800
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command21 
         Caption         =   "Back"
         Height          =   375
         Index           =   5
         Left            =   6960
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Save"
         Height          =   375
         Index           =   4
         Left            =   1440
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         Height          =   375
         Index           =   3
         Left            =   2760
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Delete"
         Height          =   375
         Index           =   2
         Left            =   4080
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Update"
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the Admission Id of the Patient to Delete or Update the Record"
      Height          =   615
      Left            =   4320
      TabIndex        =   37
      Top             =   8760
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Admission of the Patient"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   9855
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
 Public conn As New ADODB.Connection
 Public rs As New ADODB.Recordset
 Private Sub Command1_Click(Index As Integer)
 Set rs = New ADODB.Recordset
 
 
 rs.Open "select * from Admission where AdmissionId=" & Text13.Text, conn, adOpenKeyset, adLockOptimistic
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text2.Text
rs.Fields(2) = Text3.Text
rs.Fields(3) = Text4.Text
rs.Fields(4) = Check1.Value
rs.Fields(5) = Check2.Value
rs.Fields(6) = Check3.Value
rs.Fields(7) = Check4.Value
rs.Fields(8) = Text5.Text
rs.Fields(9) = Option1.Value
rs.Fields(10) = Option2.Value
rs.Fields(11) = Option3.Value
rs.Fields(12) = Option4.Value
rs.Fields(13) = Text6.Text
rs.Fields(14) = Text7.Text
rs.Fields(15) = Option5.Value
rs.Fields(16) = Option6.Value
rs.Fields(17) = Text8.Text
rs.Fields(18) = Text9.Text
rs.Fields(19) = Text10.Text
rs.Fields(20) = Text11.Text
rs.Fields(21) = Text12.Text
rs.Fields(22) = Option7.Value
rs.Fields(23) = Option8.Value
rs.Update
MsgBox "Record Updated Sucessfully"
End Sub


Private Sub Command15_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"
conn.Open
Dim a As Integer
Dim str As String
rs.Open "select max (AdmissionId)from Admission", conn, adOpenKeyset, adLockOptimistic

If IsNull(rs.Fields(0)) Then
Text1.Text = 1
Else
a = rs.Fields(0)
Text1.Text = a + 1
End If
rs.Close
End Sub

Private Sub Command2_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"
conn.Open

rs.Open "select(AdmissionId)from Admission where AdmissionId=" & Text13.Text, conn, adOpenKeyset, adLockOptimistic

If rs.EOF Then
MsgBox "No records Found.."
Else

rs.Delete
rs.Close
MsgBox ("Record Deleted Sucessfully")
End If
End Sub

Private Sub Command20_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"
conn.Open
rs.CursorLocation = adUseClient
rs.Open "select * from Admission", conn, adOpenKeyset, adLockOptimistic
rs.MoveLast

Set DataGrid1.DataSource = rs
DataGrid1.Refresh
Set rs = Nothing
End Sub

Private Sub Command21_Click(Index As Integer)
Form8.Hide
Form2.Show
End Sub

Private Sub Command22_Click(Index As Integer)

End Sub

Private Sub Command3_Click(Index As Integer)
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Text5.Text = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Text6.Text = ""
Text7.Text = ""
Option5.Value = False
Option6.Value = False
Option7.Value = False
Option8.Value = False
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
End Sub

Private Sub Command4_Click(Index As Integer)
If Text2.Text = "" Then
        MsgBox ("Name is not entered")
        
        Text2.SetFocus
        Exit Sub
    End If
    Dim num As String
num = Text4.Text
Dim validdigits As String

validdigits = "[6789]"
Dim firstdigit As String

firstdigit = Left(num, 1)
If IsNumeric(num) Then
If Len(num) = 10 And firstdigit Like validdigits Then
Else
MsgBox "Please check your Mobile Number...."
Exit Sub
End If
Else
MsgBox "enter only digits"
Text4.SetFocus
Exit Sub
End If
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from Admission ", conn, adOpenKeyset, adLockOptimistic
rs.AddNew
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text2.Text
rs.Fields(2) = Text3.Text
rs.Fields(3) = Text4.Text
rs.Fields(4) = Check1.Value
rs.Fields(5) = Check2.Value
rs.Fields(6) = Check3.Value
rs.Fields(7) = Check4.Value
rs.Fields(8) = Text5.Text
rs.Fields(9) = Option1.Value
rs.Fields(10) = Option2.Value
rs.Fields(11) = Option3.Value
rs.Fields(12) = Option4.Value
rs.Fields(13) = Text6.Text
rs.Fields(14) = Text7.Text
rs.Fields(15) = Option5.Value
rs.Fields(16) = Option6.Value
rs.Fields(17) = Text8.Text
rs.Fields(18) = Text9.Text
rs.Fields(19) = Text10.Text
rs.Fields(20) = Text11.Text
rs.Fields(21) = Text12.Text
rs.Fields(22) = Option7.Value
rs.Fields(23) = Option8.Value
rs.Update
MsgBox ("Record Added Succesfully")
rs.Close






End Sub

Private Sub Command9_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from Admission", conn, adOpenKeyset, adLockOptimistic
rs.MoveFirst

Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
Check1.Value = rs.Fields(4).Value
Check2.Value = rs.Fields(5).Value
Check3.Value = rs.Fields(6).Value
Check4.Value = rs.Fields(7).Value
Text5.Text = rs.Fields(8).Value
Option1.Value = rs.Fields(9).Value
Option2.Value = rs.Fields(10).Value
Option3.Value = rs.Fields(11).Value
Option4.Text = rs.Fields(12).Value
Text6.Text = rs.Fields(13).Value
Text7.Text = rs.Fields(14).Value
Option5.Value = rs.Fields(15).Value
Option6.Value = rs.Fields(16).Value
Text8.Text = rs.Fields(17).Value
Text9.Text = rs.Fields(18).Value
Text10.Text = rs.Fields(19).Value
Text11.Text = rs.Fields(20).Value
Text12.Text = rs.Fields(21).Value

Option7.Value = rs.Fields(22).Value
Option8.Value = rs.Fields(23).Value
Option9.Value = rs.Fields(24).Value
End Sub



Private Sub Command5_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from Admission", conn, adOpenKeyset, adLockOptimistic
rs.MoveFirst

Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
If rs.Fields(4).Value = -1 Then
Check1.Value = 1
Else
Check1.Value = 0
End If
If rs.Fields(5).Value = -1 Then
Check2.Value = 1
Else
Check2.Value = 0
End If
If rs.Fields(6).Value = -1 Then
Check3.Value = 1
Else
Check3.Value = 0
End If
If rs.Fields(7).Value = -1 Then
Check4.Value = 1
Else
Check4.Value = 0
End If



Text5.Text = rs.Fields(8).Value
If rs.Fields(9).Value = -1 Then
Option1.Value = True
Else
Option1.Value = False
End If
If rs.Fields(10).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
If rs.Fields(11).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If
If rs.Fields(12).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False
End If

Text6.Text = rs.Fields(13).Value
Text7.Text = rs.Fields(14).Value
If rs.Fields(15).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False
End If
If rs.Fields(16).Value = -1 Then
Option6.Value = True
Else
Option6.Value = False
End If

Text8.Text = rs.Fields(17).Value
Text9.Text = rs.Fields(18).Value
Text10.Text = rs.Fields(19).Value
Text11.Text = rs.Fields(20).Value
Text12.Text = rs.Fields(21).Value
If rs.Fields(22).Value = -1 Then
Option7.Value = True
Else
Option7.Value = False
End If
If rs.Fields(23).Value = -1 Then
Option8.Value = True
Else
Option8.Value = False
End If
If rs.Fields(24).Value = -1 Then
Option9.Value = True
Else
Option9.Value = False
MsgBox "You have reached to first record"

End If
End Sub


Private Sub Command6_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from Admission", conn, adOpenKeyset, adLockOptimistic

rs.MoveLast
If Not rs.EOF Then

Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
If rs.Fields(4).Value = -1 Then
Check1.Value = 1
Else
Check1.Value = 0
End If
If rs.Fields(5).Value = -1 Then
Check2.Value = 1
Else
Check2.Value = 0
End If
If rs.Fields(6).Value = -1 Then
Check3.Value = 1
Else
Check3.Value = 0
End If
If rs.Fields(7).Value = -1 Then
Check4.Value = 1
Else
Check4.Value = 0
End If



Text5.Text = rs.Fields(8).Value
If rs.Fields(9).Value = -1 Then
Option1.Value = True
Else
Option1.Value = False
End If
If rs.Fields(10).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
If rs.Fields(11).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If
If rs.Fields(12).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False
End If

Text6.Text = rs.Fields(13).Value
Text7.Text = rs.Fields(14).Value
If rs.Fields(15).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False
End If
If rs.Fields(16).Value = -1 Then
Option6.Value = True
Else
Option6.Value = False
End If

Text8.Text = rs.Fields(17).Value
Text9.Text = rs.Fields(18).Value
Text10.Text = rs.Fields(19).Value
Text11.Text = rs.Fields(20).Value
Text12.Text = rs.Fields(21).Value
If rs.Fields(22).Value = -1 Then
Option7.Value = True
Else
Option7.Value = False
End If
If rs.Fields(23).Value = -1 Then
Option8.Value = True
Else
Option8.Value = False
End If
If rs.Fields(24).Value = -1 Then
Option9.Value = True
Else
Option9.Value = False
End If
MsgBox "You have reached last record"
End If

End Sub




Private Sub Command7_Click(Index As Integer)

rs.MoveNext



If Not rs.EOF Then

Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
If rs.Fields(4).Value = -1 Then
Check1.Value = 1
Else
Check1.Value = 0
End If
If rs.Fields(5).Value = -1 Then
Check2.Value = 1
Else
Check2.Value = 0
End If
If rs.Fields(6).Value = -1 Then
Check3.Value = 1
Else
Check3.Value = 0
End If
If rs.Fields(7).Value = -1 Then
Check4.Value = 1
Else
Check4.Value = 0
End If



Text5.Text = rs.Fields(8).Value
If rs.Fields(9).Value = -1 Then
Option1.Value = True
Else
Option1.Value = False
End If
If rs.Fields(10).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
If rs.Fields(11).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If
If rs.Fields(12).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False
End If

Text6.Text = rs.Fields(13).Value
Text7.Text = rs.Fields(14).Value
If rs.Fields(15).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False
End If
If rs.Fields(16).Value = -1 Then
Option6.Value = True
Else
Option6.Value = False
End If

Text8.Text = rs.Fields(17).Value
Text9.Text = rs.Fields(18).Value
Text10.Text = rs.Fields(19).Value
Text11.Text = rs.Fields(20).Value
Text12.Text = rs.Fields(21).Value
If rs.Fields(22).Value = -1 Then
Option7.Value = True
Else
Option7.Value = False
End If
If rs.Fields(23).Value = -1 Then
Option8.Value = True
Else
Option8.Value = False
End If
If rs.Fields(24).Value = -1 Then
Option9.Value = True
Else
Option9.Value = False
End If
Else
MsgBox "You have reached last record "

End If

End Sub

Private Sub Command8_Click(Index As Integer)
rs.MovePrevious



If Not rs.BOF Then

Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
If rs.Fields(4).Value = -1 Then
Check1.Value = 1
Else
Check1.Value = 0
End If
If rs.Fields(5).Value = -1 Then
Check2.Value = 1
Else
Check2.Value = 0
End If
If rs.Fields(6).Value = -1 Then
Check3.Value = 1
Else
Check3.Value = 0
End If
If rs.Fields(7).Value = -1 Then
Check4.Value = 1
Else
Check4.Value = 0
End If



Text5.Text = rs.Fields(8).Value
If rs.Fields(9).Value = -1 Then
Option1.Value = True
Else
Option1.Value = False
End If
If rs.Fields(10).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
If rs.Fields(11).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If
If rs.Fields(12).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False
End If

Text6.Text = rs.Fields(13).Value
Text7.Text = rs.Fields(14).Value
If rs.Fields(15).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False
End If
If rs.Fields(16).Value = -1 Then
Option6.Value = True
Else
Option6.Value = False
End If

Text8.Text = rs.Fields(17).Value
Text9.Text = rs.Fields(18).Value
Text10.Text = rs.Fields(19).Value
Text11.Text = rs.Fields(20).Value
Text12.Text = rs.Fields(21).Value
If rs.Fields(22).Value = -1 Then
Option7.Value = True
Else
Option7.Value = False
End If
If rs.Fields(23).Value = -1 Then
Option8.Value = True
Else
Option8.Value = False
End If
If rs.Fields(24).Value = -1 Then
Option9.Value = True
Else
Option9.Value = False
End If
Else
MsgBox "You have reached first record "
End If
End Sub

Private Sub DataGrid1_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"
conn.Open
rs.CursorLocation = adUseClient
rs.Open "select * from Admission", conn, adOpenKeyset, adLockOptimistic
rs.MoveLast

Set DataGrid1.DataSource = rs
DataGrid1.Refresh
Set rs = Nothing
End Sub

Private Sub Form_Load()

Form8.Picture = LoadPicture("E:\Rahul Covid Project1\Resources\25.jpg")
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"

conn.Open
rs.Open "select * from Pathology", conn, adOpenKeyset, adLockOptimistic



End Sub

