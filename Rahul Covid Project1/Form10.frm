VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   8025
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   LinkTopic       =   "Form10"
   Picture         =   "Form10.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   13560
      TabIndex        =   50
      Top             =   2520
      Width           =   2055
      Begin VB.OptionButton Option5 
         Caption         =   "Dr.Kolap"
         Height          =   195
         Left            =   600
         TabIndex        =   53
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Dr. Dasari"
         Height          =   375
         Left            =   1200
         TabIndex        =   52
         Top             =   120
         Width           =   855
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Dr Nair"
         Height          =   375
         Left            =   120
         TabIndex        =   51
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.TextBox Text14 
      Height          =   525
      Left            =   13440
      TabIndex        =   49
      Top             =   6720
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Menus"
      Height          =   855
      Left            =   0
      TabIndex        =   36
      Top             =   7560
      Width           =   15495
      Begin VB.CommandButton Command11 
         Caption         =   "Generate Report"
         Height          =   375
         Left            =   14280
         TabIndex        =   56
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         Caption         =   "View Database"
         Height          =   375
         Left            =   12840
         TabIndex        =   55
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Delete 
         Caption         =   "Delete"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   44
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Clear 
         Caption         =   "Clear"
         Height          =   375
         Index           =   3
         Left            =   1680
         TabIndex        =   43
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Save"
         Height          =   375
         Index           =   4
         Left            =   3240
         TabIndex        =   42
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Back"
         Height          =   375
         Index           =   5
         Left            =   4920
         TabIndex        =   41
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Move First"
         Height          =   375
         Index           =   0
         Left            =   6480
         TabIndex        =   40
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Move Last"
         Height          =   375
         Index           =   1
         Left            =   11280
         TabIndex        =   39
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Move Next"
         Height          =   375
         Index           =   2
         Left            =   8040
         TabIndex        =   38
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Move Previous"
         Height          =   375
         Index           =   3
         Left            =   9600
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Access Admission  Details"
      Height          =   855
      Left            =   0
      TabIndex        =   15
      Top             =   6480
      Width           =   7095
      Begin VB.CommandButton Command1 
         Caption         =   "MoveFirst"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "MoveLast"
         Height          =   375
         Index           =   2
         Left            =   5400
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Movenext"
         Height          =   375
         Index           =   3
         Left            =   1800
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "MovePrevious"
         Height          =   375
         Index           =   5
         Left            =   3720
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Personal Details of the Patient"
      Height          =   4935
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   15735
      Begin VB.TextBox Expenses 
         Height          =   405
         Left            =   13560
         TabIndex        =   48
         Text            =   "0"
         Top             =   3840
         Width           =   2055
      End
      Begin VB.TextBox Text13 
         Enabled         =   0   'False
         Height          =   405
         Left            =   13680
         TabIndex        =   46
         Top             =   3000
         Width           =   1935
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   13560
         TabIndex        =   35
         Top             =   4440
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "No"
         Height          =   255
         Left            =   14640
         TabIndex        =   34
         Top             =   2520
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Yes"
         Height          =   375
         Left            =   13560
         TabIndex        =   33
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox Text11 
         Height          =   405
         Left            =   13680
         TabIndex        =   32
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text10 
         Height          =   405
         Left            =   8520
         TabIndex        =   31
         Top             =   3720
         Width           =   2175
      End
      Begin VB.TextBox Text9 
         Height          =   405
         Left            =   8520
         TabIndex        =   30
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox Text8 
         Height          =   495
         Left            =   8520
         TabIndex        =   29
         Top             =   2160
         Width           =   2175
      End
      Begin VB.TextBox Text7 
         Height          =   405
         Left            =   8520
         TabIndex        =   28
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   8520
         TabIndex        =   27
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2880
         TabIndex        =   26
         Top             =   3960
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   2880
         TabIndex        =   25
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2880
         TabIndex        =   24
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   2880
         TabIndex        =   23
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2880
         TabIndex        =   22
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Insurance Amount"
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
         Left            =   11040
         TabIndex        =   47
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Final"
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
         Left            =   11040
         TabIndex        =   45
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Addional"
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
         Index           =   14
         Left            =   11040
         TabIndex        =   21
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Insurance"
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
         Left            =   11040
         TabIndex        =   14
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Discharge Issued By"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   11
         Left            =   11040
         TabIndex        =   13
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bed Allocated"
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
         Index           =   10
         Left            =   11040
         TabIndex        =   12
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Food Charges"
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
         Index           =   6
         Left            =   5640
         TabIndex        =   11
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Medicine Charges"
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
         Index           =   5
         Left            =   5640
         TabIndex        =   10
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hospital Charges"
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
         Index           =   7
         Left            =   5640
         TabIndex        =   9
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Number of Days "
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
         Index           =   8
         Left            =   5640
         TabIndex        =   8
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Date of Discharge"
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
         Left            =   5640
         TabIndex        =   7
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Discharge ID"
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
         Index           =   9
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   495
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   495
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Date of Neg Report Generated By Lab"
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
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Date Of Admission"
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
         Index           =   3
         Left            =   120
         TabIndex        =   2
         Top             =   3960
         Width           =   2415
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4215
      Left            =   0
      TabIndex        =   20
      Top             =   8640
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   7435
      _Version        =   393216
      BackColor       =   14737632
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
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enter the Id to Update or Delete the Record"
      Height          =   615
      Left            =   8640
      TabIndex        =   54
      Top             =   6720
      Width           =   3735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Issue Discharge Certificate to the Patient"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   240
      Width           =   6855
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public conn2 As New ADODB.Connection
 Public conn As New ADODB.Connection
 Public conn1 As New ADODB.Connection
  Public rs1 As New ADODB.Recordset
 Public rs As New ADODB.Recordset
 Public rs2 As New ADODB.Recordset

Private Sub Clear_Click(Index As Integer)
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Text12.Text = ""
Expenses.Text = ""
Text13.Text = ""
End Sub

Private Sub Command1_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"

conn.Open
rs.Open "select * from Admission", conn, adOpenKeyset, adLockOptimistic

rs.MoveFirst
MsgBox "You have reached First record"
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(8).Value

Text5.Text = rs.Fields(17).Value

Text11.Text = rs.Fields(21).Value

Set conn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset

rs1.Open "select * from Treatment", conn, adOpenKeyset, adLockOptimistic

rs1.MoveFirst
Text7.Text = rs1.Fields(8).Value
Text8.Text = rs1.Fields(9).Value
Text9.Text = rs1.Fields(10).Value
Text10.Text = rs1.Fields(11).Value


End Sub

Private Sub Command10_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"

conn.Open
rs.CursorLocation = adUseClient
rs.Open "select * from Discharge", conn, adOpenKeyset, adLockOptimistic
rs.MoveLast

Set DataGrid1.DataSource = rs
DataGrid1.Refresh
Set rs = Nothing
End Sub

Private Sub Command11_Click()
Dim a As Integer
a = InputBox("Enter the Id")
DataEnvironment1.Command1 (a)
DataReport4.Show
DataReport4.Refresh
DataEnvironment1.rsCommand1.Close
End Sub

Private Sub Command2_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"

conn2.Open
rs.Open "select * from Admission", conn, adOpenKeyset, adLockOptimistic

rs.MoveLast
If Not rs.EOF Then
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(8).Value

Text5.Text = rs.Fields(17).Value

Text11.Text = rs.Fields(21).Value
Else
MsgBox "You have reached last record"
End If

Set conn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset

rs1.Open "select * from Treatment", conn, adOpenKeyset, adLockOptimistic

rs1.MoveLast

If Not rs1.EOF Then
Text7.Text = rs1.Fields(8).Value
Text8.Text = rs1.Fields(9).Value
Text9.Text = rs1.Fields(10).Value
Text10.Text = rs1.Fields(11).Value
Else
MsgBox "You have reached last record"
End If
End Sub

Private Sub Command3_Click(Index As Integer)
rs.MoveNext
If Not rs.EOF Then
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(8).Value

Text5.Text = rs.Fields(17).Value

Text11.Text = rs.Fields(21).Value
End If
rs1.MoveNext
If Not rs1.EOF Then
Text7.Text = rs1.Fields(8).Value
Text8.Text = rs1.Fields(9).Value
Text9.Text = rs1.Fields(10).Value
Text10.Text = rs1.Fields(11).Value

Else
MsgBox "You have reached last record"
End If



End Sub

Private Sub Command4_Click(Index As Integer)
rs.MovePrevious
If Not rs.BOF Then
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(8).Value

Text5.Text = rs.Fields(17).Value

Text11.Text = rs.Fields(21).Value
End If
rs1.MovePrevious
If Not rs1.BOF Then
Text7.Text = rs1.Fields(8).Value
Text8.Text = rs1.Fields(9).Value
Text9.Text = rs1.Fields(10).Value
Text10.Text = rs1.Fields(11).Value
Text10.Text = rs1.Fields(12).Value
Else
 MsgBox "You have reached first record"

End If
End Sub

Private Sub Command5_Click(Index As Integer)
If Text2.Text = "" Then
        MsgBox ("Name is not entered"), vbInformation
        
        Text2.SetFocus
        Exit Sub
    End If
   

Set conn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset
conn2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"

conn2.Open
rs2.Open "select * from Discharge", conn, adOpenKeyset, adLockOptimistic
rs2.AddNew
rs2.Fields(0) = Text1.Text
rs2.Fields(1) = Text2.Text
rs2.Fields(2) = Text3.Text
rs2.Fields(3) = Text4.Text
rs2.Fields(4) = Text5.Text
rs2.Fields(5) = Text6.Text
rs2.Fields(6) = Text7.Text
rs2.Fields(7) = Text8.Text
rs2.Fields(8) = Text9.Text
rs2.Fields(9) = Text10.Text
rs2.Fields(10) = Text11.Text
rs2.Fields(11) = Option3.Value
rs2.Fields(12) = Option4.Value
rs2.Fields(13) = Option5.Value
rs2.Fields(14) = Option1.Value
rs2.Fields(15) = Option2.Value
rs2.Fields(16) = Text13.Text
rs2.Fields(17) = Expenses.Text
rs2.Fields(18) = Text12.Text
MsgBox ("Record Added Succesfully")


rs2.Update
End Sub

Private Sub Command6_Click(Index As Integer)
Set conn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset
conn2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"

conn2.Open
rs2.Open "select * from Discharge", conn, adOpenKeyset, adLockOptimistic
MsgBox "You have reached last record", vbInformation
rs2.MoveLast
If Not rs2.EOF Then
Text1.Text = rs2.Fields(0).Value
Text2.Text = rs2.Fields(1).Value
Text3.Text = rs2.Fields(2).Value
Text4.Text = rs2.Fields(3).Value
Text5.Text = rs2.Fields(4).Value
Text6.Text = rs2.Fields(5).Value
Text7.Text = rs2.Fields(6).Value
Text8.Text = rs2.Fields(7).Value
Text9.Text = rs2.Fields(8).Value
Text10.Text = rs2.Fields(9).Value
Text11.Text = rs2.Fields(10).Value

If rs2.Fields(11).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If
If rs2.Fields(12).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False
End If
If rs2.Fields(13).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False
End If
If rs2.Fields(14).Value = -1 Then
Option1.Value = True
Else
Option1.Value = False
End If
If rs2.Fields(15).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
Text13.Text = rs2.Fields(16).Value
Expenses.Text = rs2.Fields(17).Value
Text12.Text = rs2.Fields(18).Value
End If
End Sub

Private Sub Command7_Click(Index As Integer)
rs2.MoveNext
If Not rs2.EOF Then
Text1.Text = rs2.Fields(0).Value
Text2.Text = rs2.Fields(1).Value
Text3.Text = rs2.Fields(2).Value
Text4.Text = rs2.Fields(3).Value
Text5.Text = rs2.Fields(4).Value
Text6.Text = rs2.Fields(5).Value
Text7.Text = rs2.Fields(6).Value
Text8.Text = rs2.Fields(7).Value
Text9.Text = rs2.Fields(8).Value
Text10.Text = rs2.Fields(9).Value
Text11.Text = rs2.Fields(10).Value

If rs2.Fields(11).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If
If rs2.Fields(12).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False
End If
If rs2.Fields(13).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False
End If
If rs2.Fields(14).Value = -1 Then
Option1.Value = True
Else
Option1.Value = False
End If
If rs2.Fields(15).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
Text13.Text = rs2.Fields(16).Value
Expenses.Text = rs2.Fields(17).Value
Text12.Text = rs2.Fields(18).Value
Else
MsgBox "You have reached last record", vbInformation


End If

End Sub

Private Sub Command8_Click(Index As Integer)
rs2.MovePrevious
If Not rs2.BOF Then
Text1.Text = rs2.Fields(0).Value
Text2.Text = rs2.Fields(1).Value
Text3.Text = rs2.Fields(2).Value
Text4.Text = rs2.Fields(3).Value
Text5.Text = rs2.Fields(4).Value
Text6.Text = rs2.Fields(5).Value
Text7.Text = rs2.Fields(6).Value
Text8.Text = rs2.Fields(7).Value
Text9.Text = rs2.Fields(8).Value
Text10.Text = rs2.Fields(9).Value
Text11.Text = rs2.Fields(10).Value

If rs2.Fields(11).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If
If rs2.Fields(12).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False
End If
If rs2.Fields(13).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False
End If
If rs2.Fields(14).Value = -1 Then
Option1.Value = True
Else
Option1.Value = False
End If
If rs2.Fields(15).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
Text13.Text = rs2.Fields(16).Value
Expenses.Text = rs2.Fields(17).Value
Text12.Text = rs2.Fields(18).Value
Else
MsgBox "You have reached first record", vbInformation

End If
End Sub

Private Sub Command9_Click(Index As Integer)
Set conn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset
conn2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"

conn2.Open
rs2.Open "select * from Discharge", conn, adOpenKeyset, adLockOptimistic

rs2.MoveFirst
MsgBox "You have reached first record ", vbInformation
Text1.Text = rs2.Fields(0).Value
Text2.Text = rs2.Fields(1).Value
Text3.Text = rs2.Fields(2).Value
Text4.Text = rs2.Fields(3).Value
Text5.Text = rs2.Fields(4).Value
Text6.Text = rs2.Fields(5).Value
Text7.Text = rs2.Fields(6).Value
Text8.Text = rs2.Fields(7).Value
Text9.Text = rs2.Fields(8).Value
Text10.Text = rs2.Fields(9).Value
Text11.Text = rs2.Fields(10).Value

If rs2.Fields(11).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If
If rs2.Fields(12).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False
End If
If rs2.Fields(13).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False
End If
If rs2.Fields(14).Value = -1 Then
Option1.Value = True
Else
Option1.Value = False
End If
If rs2.Fields(15).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
Text13.Text = rs2.Fields(16).Value
Expenses.Text = rs2.Fields(17).Value
Text12.Text = rs2.Fields(18).Value



End Sub

Private Sub Delete_Click(Index As Integer)
Set conn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset
conn2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn2.Open

rs2.Open "select(PatientId)from Pathology where PatientId=" & Text14.Text, conn, adOpenKeyset, adLockOptimistic

If rs2.EOF = True Then
MsgBox "Record not found", vbInformation
Else
rs2.Delete
rs2.Close
MsgBox ("Record Deleted Sucessfully")
End If
End Sub

Private Sub Form_Load()
Form10.Picture = LoadPicture("E:\Rahul Covid Project1\Resources\55.jpg")
Set conn = New ADODB.Connection
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"

conn.Open
rs.Open "select * from Admission", conn, adOpenKeyset, adLockOptimistic
rs.MoveLast
Set conn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset

rs1.Open "select * from Treatment", conn, adOpenKeyset, adLockOptimistic
rs1.MoveLast
Set conn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset

rs2.Open "select * from Treatment", conn, adOpenKeyset, adLockOptimistic
rs2.MoveLast
End Sub

Private Sub Option1_Click()
Dim s As Long
Text13.Enabled = True
Dim l As Long
l = CLng(InputBox("Enter the Amount to Redeem..."))
Text13.Text = l
s = CLng(Text7.Text) + CLng(Text8.Text) + CLng(Text9.Text) + CLng(Text10.Text) + CLng(Expenses.Text) - CLng(Text13.Text)
Text12.Text = CLng(s)


End Sub

Private Sub Option2_Click()

Dim s As Long
Text13.Text = 0
s = CLng(Text7.Text) + CLng(Text8.Text) + CLng(Text9.Text) + CLng(Text10.Text) + CLng(Expenses.Text)
Text12.Text = CLng(s)
End Sub

