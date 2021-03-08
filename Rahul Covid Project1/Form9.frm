VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form9 
   Caption         =   "Form9"
   ClientHeight    =   8445
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   LinkTopic       =   "Form9"
   Picture         =   "Form9.frx":0000
   ScaleHeight     =   8445
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text9 
      Height          =   495
      Left            =   11640
      TabIndex        =   43
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Menus "
      Height          =   855
      Left            =   0
      TabIndex        =   35
      Top             =   7800
      Width           =   15015
      Begin VB.CommandButton Command7 
         Caption         =   "View Database"
         Height          =   375
         Left            =   13560
         TabIndex        =   46
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Clear"
         Height          =   375
         Left            =   10680
         TabIndex        =   44
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Delete"
         Height          =   375
         Left            =   9240
         TabIndex        =   42
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   7320
         TabIndex        =   41
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton MoveFirst 
         Caption         =   "MoveFirst"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   40
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton MoveLast 
         Cancel          =   -1  'True
         Caption         =   "MoveLast"
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   39
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton MoveNext 
         Caption         =   "MoveNext"
         Height          =   375
         Index           =   0
         Left            =   3960
         TabIndex        =   38
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton MovePrevious 
         Caption         =   "MovePrevious"
         Height          =   375
         Index           =   0
         Left            =   5640
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command44 
         Caption         =   "Back"
         Height          =   375
         Index           =   1
         Left            =   12120
         TabIndex        =   36
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Access Admission Details"
      Height          =   855
      Left            =   0
      TabIndex        =   13
      Top             =   6720
      Width           =   7095
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   375
         Index           =   5
         Left            =   12120
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "MovePrevious"
         Height          =   375
         Index           =   3
         Left            =   3840
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "MoveNext"
         Height          =   375
         Index           =   2
         Left            =   2040
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "MoveLast"
         Height          =   375
         Index           =   1
         Left            =   5640
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "MoveFirst"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Treatement and the Services taken by the Patient"
      Height          =   4935
      Index           =   1
      Left            =   8400
      TabIndex        =   7
      Top             =   1320
      Width           =   6855
      Begin VB.TextBox Text8 
         Height          =   405
         Left            =   3120
         TabIndex        =   34
         Top             =   3840
         Width           =   2175
      End
      Begin VB.TextBox Text7 
         Height          =   405
         Left            =   3120
         TabIndex        =   33
         Top             =   3000
         Width           =   2175
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   3120
         TabIndex        =   32
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   3120
         TabIndex        =   31
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   3120
         TabIndex        =   30
         Top             =   720
         Width           =   2175
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
         Left            =   120
         TabIndex        =   12
         Top             =   720
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
         Left            =   120
         TabIndex        =   11
         Top             =   1440
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
         Left            =   120
         TabIndex        =   10
         Top             =   2280
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
         Left            =   120
         TabIndex        =   9
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total Expenses"
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
         Left            =   120
         TabIndex        =   8
         Top             =   3840
         Width           =   2415
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Personal Details of the Patient"
      Height          =   4935
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   7575
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   3120
         TabIndex        =   29
         Top             =   2160
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3120
         TabIndex        =   28
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   3120
         TabIndex        =   27
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Frame Frame7 
         Height          =   1215
         Left            =   3120
         TabIndex        =   21
         Top             =   2640
         Width           =   2415
         Begin VB.OptionButton Option1 
            Caption         =   "General"
            Height          =   195
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "ICU"
            Height          =   435
            Left            =   1320
            TabIndex        =   23
            Top             =   120
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Special"
            Height          =   255
            Left            =   720
            TabIndex        =   22
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Index           =   0
         Left            =   3000
         TabIndex        =   20
         Top             =   3960
         Width           =   3135
         Begin VB.OptionButton Option5 
            Caption         =   "Pos(Mi)"
            Height          =   495
            Left            =   1200
            TabIndex        =   26
            Top             =   240
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Pos(Mj)"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Status"
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
         TabIndex        =   6
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ward Allocated"
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
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bed"
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
         TabIndex        =   3
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Patient ID"
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
         TabIndex        =   2
         Top             =   720
         Width           =   2415
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4215
      Left            =   0
      TabIndex        =   19
      Top             =   8880
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   7435
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
   Begin VB.Label Label3 
      Caption         =   "Enter the Patient Id to Update or Delete the Record"
      Height          =   855
      Left            =   8520
      TabIndex        =   45
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Treatment of the Patient"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   0
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Option Explicit
 Public conn As New ADODB.Connection
 Public rs As New ADODB.Recordset
 Public conn1 As New ADODB.Connection
 Public rs1 As New ADODB.Recordset

Private Sub cmdSave_Click()
If Text2.Text = "" Then
        MsgBox ("Name is not entered")
        
        Text2.SetFocus
        Exit Sub
    End If
   

Set conn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset
conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"
conn1.Open
rs1.Open "select * from Treatment", conn1, adOpenKeyset, adLockOptimistic
rs1.AddNew
rs1.Fields(0).Value = Text1.Text
rs1.Fields(1).Value = Text2.Text
rs1.Fields(2).Value = Text3.Text
rs1.Fields(3).Value = Option1.Value
rs1.Fields(4).Value = Option2.Value
rs1.Fields(5).Value = Option3.Value
rs1.Fields(6).Value = Option4.Value
rs1.Fields(7).Value = Option5.Value

rs1.Fields(8).Value = Text4.Text

rs1.Fields(9).Value = Text5.Text
rs1.Fields(10).Value = Text6.Text
rs1.Fields(11).Value = Text7.Text
rs1.Fields(12).Value = Text8.Text
rs1.Update
MsgBox ("Record Added Successfully")

End Sub

 Private Sub Command1_Click(Index As Integer)

Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from Admission", conn, adOpenKeyset, adLockOptimistic

rs.MoveFirst
MsgBox "You have reached First record"
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(20).Value
If rs.Fields(21).Value = -1 Then
Option1.Value = True
Else
Option2.Value = False
End If
If rs.Fields(22).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
If rs.Fields(23).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If

If rs.Fields(11).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False
End If


If rs.Fields(12).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False
End If


End Sub

Private Sub Command2_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from Admission", conn, adOpenKeyset, adLockOptimistic

rs.MoveLast
If Not rs.EOF Then
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(20).Value
If rs.Fields(21).Value = -1 Then
Option1.Value = True
Else
Option2.Value = False
End If
If rs.Fields(22).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
If rs.Fields(23).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If

If rs.Fields(11).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False
End If


If rs.Fields(12).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False
End If
Else
MsgBox "You have reached first record .."
End If
End Sub

Private Sub Command3_Click(Index As Integer)
rs.MoveNext




If Not rs.EOF Then

Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(20).Value
If rs.Fields(21).Value = -1 Then
Option1.Value = True
Else
Option2.Value = False
End If
If rs.Fields(22).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
If rs.Fields(23).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If

If rs.Fields(11).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False
End If


If rs.Fields(12).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False
End If
Else
MsgBox "You have reached First record"
End If
End Sub

Private Sub Command4_Click(Index As Integer)

rs.MovePrevious
If Not rs.BOF Then

Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(20).Value
If rs.Fields(21).Value = -1 Then
Option1.Value = True
Else
Option2.Value = False
End If
If rs.Fields(22).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
If rs.Fields(23).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If

If rs.Fields(11).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False
End If


If rs.Fields(12).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False
End If
Else
MsgBox "You have reached first record of Admission Table.."
End If
End Sub

Private Sub Command44_Click(Index As Integer)
Form9.Hide
Form2.Show
End Sub

Private Sub Command5_Click()
Set conn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset
conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn1.Open

rs1.Open "select(AdmissionId)from Treatment where AdmissionId=" & Text9.Text, conn, adOpenKeyset, adLockOptimistic

If rs1.EOF = True Then
MsgBox "No more records.."
Else
rs1.Delete
rs1.Close
MsgBox ("Record Deleted Sucessfully")
End If
End Sub

Private Sub Command6_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""

End Sub

Private Sub Command7_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.CursorLocation = adUseClient
rs.Open "select * from Treatment", conn, adOpenKeyset, adLockOptimistic
rs.MoveLast

Set DataGrid1.DataSource = rs
DataGrid1.Refresh
Set rs = Nothing
End Sub

Private Sub Form_Load()
Form9.Picture = LoadPicture("C:\Rahul Covid Project1\Resources\wp2386761.jpg")
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from Admission", conn, adOpenKeyset, adLockOptimistic
rs.MoveLast

Set conn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset
conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn1.Open
rs1.Open "select * from Treatment", conn, adOpenKeyset
End Sub

Private Sub MoveFirst_Click(Index As Integer)
Set conn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset
conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn1.Open
rs1.Open "select * from Treatment", conn, adOpenKeyset, adLockOptimistic

rs1.MoveFirst
MsgBox "You have reached first record of test table...#NoMoreRecords.. "
Text1.Text = rs1.Fields(0).Value
Text2.Text = rs1.Fields(1).Value
Text3.Text = rs1.Fields(2).Value
If rs1.Fields(3).Value = -1 Then
Option1.Value = True
Else
Option2.Value = False
End If
If rs1.Fields(4).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
If rs1.Fields(5).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If

If rs1.Fields(6).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False
End If


If rs1.Fields(7).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False
End If
Text4.Text = rs1.Fields(8).Value
Text5.Text = rs1.Fields(9).Value
Text6.Text = rs1.Fields(10).Value
Text7.Text = rs1.Fields(11).Value
Text8.Text = rs1.Fields(12).Value


End Sub

Private Sub MoveLast_Click(Index As Integer)
Set conn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset
conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn1.Open
rs1.Open "select * from Treatment", conn, adOpenKeyset, adLockOptimistic

rs1.MoveLast
MsgBox "You have reached last record of test table...#NoMoreRecords.. "
If Not rs1.EOF Then
Text1.Text = rs1.Fields(0).Value
Text2.Text = rs1.Fields(1).Value
Text3.Text = rs1.Fields(2).Value
If rs1.Fields(3).Value = -1 Then
Option1.Value = True
Else
Option2.Value = False
End If
If rs1.Fields(4).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
If rs1.Fields(5).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If

If rs1.Fields(6).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False
End If


If rs1.Fields(7).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False
End If
Text4.Text = rs1.Fields(8).Value
Text5.Text = rs1.Fields(9).Value
Text6.Text = rs1.Fields(10).Value
Text7.Text = rs1.Fields(11).Value
Text8.Text = rs1.Fields(12).Value
Else

End If


End Sub

Private Sub MoveNext_Click(Index As Integer)
rs1.MoveNext
If Not rs1.EOF Then
Text1.Text = rs1.Fields(0).Value
Text2.Text = rs1.Fields(1).Value
Text3.Text = rs1.Fields(2).Value
If rs1.Fields(3).Value = -1 Then
Option1.Value = True
Else
Option2.Value = False
End If
If rs1.Fields(4).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
If rs1.Fields(5).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If

If rs1.Fields(6).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False
End If


If rs1.Fields(7).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False
End If
Text4.Text = rs1.Fields(8).Value
Text5.Text = rs1.Fields(9).Value
Text6.Text = rs1.Fields(10).Value
Text7.Text = rs1.Fields(11).Value
Text8.Text = rs1.Fields(12).Value
Else

MsgBox "No more records"
End If



End Sub

Private Sub MovePrevious_Click(Index As Integer)
rs1.MovePrevious
If Not rs1.BOF Then

Text1.Text = rs1.Fields(0).Value
Text2.Text = rs1.Fields(1).Value
Text3.Text = rs1.Fields(2).Value
If rs1.Fields(3).Value = -1 Then
Option1.Value = True
Else
Option2.Value = False
End If
If rs1.Fields(4).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
If rs1.Fields(5).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If

If rs1.Fields(6).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False
End If


If rs1.Fields(7).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False
End If
Text4.Text = rs1.Fields(8).Value
Text5.Text = rs1.Fields(9).Value
Text6.Text = rs1.Fields(10).Value
Text7.Text = rs1.Fields(11).Value
Text8.Text = rs1.Fields(12).Value
Else

MsgBox "Record Not found"
End If
End Sub

Private Sub Text4_Change()
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
a = Text4.Text
For b = 1 To a
d = 300 + d
Next
Text5.Text = d
End Sub
