VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   7890
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   7890
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command20 
      Caption         =   "DataGrid Refresh"
      Height          =   375
      Index           =   0
      Left            =   13680
      TabIndex        =   41
      Top             =   10200
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   405
      Left            =   11400
      TabIndex        =   34
      Top             =   9240
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Menus for Generating Reports "
      Height          =   855
      Index           =   0
      Left            =   0
      TabIndex        =   23
      Top             =   9840
      Width           =   15855
      Begin VB.CommandButton Command2 
         Caption         =   "Update"
         Height          =   375
         Index           =   1
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton CommandDel3 
         Caption         =   "Delete"
         Height          =   375
         Index           =   2
         Left            =   1680
         TabIndex        =   31
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear"
         Height          =   375
         Index           =   3
         Left            =   3240
         TabIndex        =   30
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Save"
         Height          =   375
         Index           =   4
         Left            =   4800
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command19 
         Caption         =   "Back"
         Height          =   375
         Index           =   5
         Left            =   6360
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Move First"
         Height          =   375
         Index           =   0
         Left            =   7920
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Move Last"
         Height          =   375
         Index           =   1
         Left            =   12240
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Move Next"
         Height          =   375
         Index           =   2
         Left            =   9480
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Move Previous"
         Height          =   375
         Index           =   3
         Left            =   10920
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Accesing Details of  Pathology Details"
      Height          =   855
      Index           =   1
      Left            =   7440
      TabIndex        =   11
      Top             =   7080
      Width           =   8655
      Begin VB.CommandButton Command1 
         Caption         =   "Cancel"
         Height          =   375
         Index           =   1
         Left            =   6000
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton MoveLast 
         Caption         =   "MoveLast"
         Height          =   375
         Index           =   0
         Left            =   4560
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton MoveFirst 
         Caption         =   "Move First"
         Height          =   435
         Index           =   5
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Movenext 
         Caption         =   "MoveNext"
         Height          =   375
         Index           =   4
         Left            =   1560
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton MovePrevious 
         Caption         =   "MovePrevious"
         Height          =   375
         Index           =   3
         Left            =   2880
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   4575
      Left            =   7440
      TabIndex        =   10
      Top             =   1680
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   8070
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
   Begin VB.Frame Frame3 
      Height          =   8055
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   7095
      Begin VB.OptionButton Option5 
         Caption         =   "No"
         Height          =   495
         Left            =   4440
         TabIndex        =   39
         Top             =   6600
         Width           =   975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Yes"
         Height          =   255
         Left            =   3240
         TabIndex        =   38
         Top             =   6720
         Width           =   855
      End
      Begin VB.Frame Frame5 
         Height          =   855
         Index           =   0
         Left            =   2880
         TabIndex        =   33
         Top             =   3600
         Width           =   3135
         Begin VB.OptionButton Option3 
            Caption         =   "Pos(Mi)"
            Height          =   495
            Left            =   2160
            TabIndex        =   37
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Pos(Mj)"
            Height          =   195
            Left            =   1200
            TabIndex        =   36
            Top             =   360
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Neg"
            Height          =   375
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   3120
         TabIndex        =   22
         Top             =   5880
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3120
         TabIndex        =   21
         Top             =   5160
         Width           =   2175
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3120
         TabIndex        =   20
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3120
         TabIndex        =   19
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3120
         TabIndex        =   18
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   3120
         TabIndex        =   17
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Left            =   240
         TabIndex        =   9
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   495
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Adress"
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
         Left            =   240
         TabIndex        =   7
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   495
         Index           =   14
         Left            =   120
         TabIndex        =   6
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Result"
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
         Index           =   15
         Left            =   120
         TabIndex        =   5
         Top             =   3720
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
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
         Index           =   7
         Left            =   120
         TabIndex        =   4
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Receipt Issued"
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
         Left            =   120
         TabIndex        =   3
         Top             =   5760
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Advice for Home Quarantine"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   4
         Left            =   120
         TabIndex        =   2
         Top             =   6600
         Width           =   2415
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Enter the Patients Id to Delete or Update the Record"
      Height          =   495
      Left            =   7800
      TabIndex        =   40
      Top             =   9240
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "   Issue Receipt for the Report generated by Pathology Lab"
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
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   10935
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Public conn As New ADODB.Connection
  Public conn1 As New ADODB.Connection
 Public rs As New ADODB.Recordset
  Public rs1 As New ADODB.Recordset
 
Private Sub Form4_Load()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from Pathology", conn, adOpenKeyset, adLockOptimistic
rs.MoveLast
End Sub



Private Sub Command19_Click(Index As Integer)
Form4.Hide
Form2.Show
End Sub

Private Sub Command2_Click(Index As Integer)
'Set conn = New ADODB.Connection
Set rs1 = New ADODB.Recordset
'conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"



rs1.Open "select * from Test where PatientId=" & Text7.Text, conn, adOpenKeyset, adLockOptimistic
If rs1.EOF = True Then
MsgBox "No record Found.."
Else
rs1.Fields(0).Value = Text1.Text
rs1.Fields(1).Value = Text2.Text
rs1.Fields(2).Value = Text3.Text
rs1.Fields(3).Value = Text4.Text

rs1.Fields(4).Value = Option1.Value
rs1.Fields(5).Value = Option2.Value
rs1.Fields(6).Value = Option3.Value
rs1.Fields(9).Value = Text5.Text
rs1.Fields(10).Value = Text6.Text
rs1.Fields(7).Value = Option4.Value
rs1.Fields(8).Value = Option5.Value
rs1.Update
MsgBox ("Record Updated Succesfully")
rs1.Close
End If
End Sub

Private Sub Command20_Click(Index As Integer)

Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.CursorLocation = adUseClient
rs.Open "select * from Test", conn, adOpenKeyset, adLockOptimistic
rs.MoveLast

Set DataGrid1.DataSource = rs
DataGrid1.Refresh
Set rs = Nothing

End Sub

Private Sub Command4_Click(Index As Integer)
Unload Me

Form2.Show
End Sub

Private Sub Command5_Click(Index As Integer)
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
Dim i As Integer
Set conn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset
conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn1.Open
rs1.Open "select * from Test", conn1, adOpenKeyset, adLockOptimistic
rs1.AddNew
rs1.Fields(0) = Text1.Text
rs1.Fields(1) = Text2.Text
rs1.Fields(2) = Text3.Text
rs1.Fields(3) = Text4.Text

rs1.Fields(4) = Option1.Value
rs1.Fields(5) = Option2.Value
rs1.Fields(6) = Option3.Value
rs1.Fields(9) = Text5.Text
rs1.Fields(10) = Text6.Text
rs1.Fields(7) = Option4.Value
rs1.Fields(8) = Option5.Value
rs1.Update
MsgBox ("Record Added Succesfully")
rs1.Close

End Sub

Private Sub Command6_Click(Index As Integer)
Set conn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset
conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn1.Open
rs1.Open "select * from test", conn, adOpenKeyset, adLockOptimistic

rs1.MoveLast
If Not rs1.EOF Then
Text1.Text = rs1.Fields(0).Value
Text2.Text = rs1.Fields(1).Value
Text3.Text = rs1.Fields(2).Value
Text4.Text = rs1.Fields(3).Value

If rs.Fields(5).Value = -1 Then
Option1.Value = True

Else
Option1.Value = False
End If
If rs.Fields(6).Value = -1 Then
Option2.Value = True

Else
Option2.Value = False
End If
If rs.Fields(7).Value = -1 Then
Option3.Value = True

Else
Option3.Value = False



End If
If rs.Fields(8).Value = -1 Then
Option4.Value = True

Else
Option4.Value = False
End If
If rs.Fields(9).Value = -1 Then
Option5.Value = True

Else
Option5.Value = False
End If
Text5.Text = rs.Fields(10).Value
Text6.Text = rs.Fields(11).Value
Else
MsgBox "You have reached to the last record of Test Table"
End If


End Sub

Private Sub Command7_Click(Index As Integer)
rs1.MoveNext
If Not rs1.EOF Then
Text1.Text = rs1.Fields(0).Value
Text2.Text = rs1.Fields(1).Value
Text3.Text = rs1.Fields(2).Value
Text4.Text = rs1.Fields(3).Value

If rs1.Fields(4).Value = -1 Then
Option1.Value = True

Else
Option1.Value = False
End If
If rs1.Fields(5).Value = -1 Then
Option2.Value = True

Else
Option2.Value = False
End If
If rs1.Fields(6).Value = -1 Then
Option3.Value = True

Else
Option3.Value = False



End If
If rs1.Fields(7).Value = -1 Then
Option4.Value = True

Else
Option4.Value = False
End If
If rs1.Fields(8).Value = -1 Then
Option5.Value = True

Else
Option5.Value = False
End If
Text5.Text = rs1.Fields(9).Value
Text6.Text = rs1.Fields(10).Value
Else
MsgBox "You have reached to the last record of Test Table"

End If

End Sub

Private Sub Command8_Click(Index As Integer)
rs1.MovePrevious
If Not rs1.BOF Then
Text1.Text = rs1.Fields(0).Value
Text2.Text = rs1.Fields(1).Value
Text3.Text = rs1.Fields(2).Value
Text4.Text = rs1.Fields(3).Value

If rs1.Fields(4).Value = -1 Then
Option1.Value = True

Else
Option1.Value = False
End If
If rs1.Fields(5).Value = -1 Then
Option2.Value = True

Else
Option2.Value = False
End If
If rs1.Fields(6).Value = -1 Then
Option3.Value = True

Else
Option3.Value = False



End If
If rs1.Fields(7).Value = -1 Then
Option4.Value = True

Else
Option4.Value = False
End If
If rs1.Fields(8).Value = -1 Then
Option5.Value = True

Else
Option5.Value = False
End If
Text5.Text = rs1.Fields(9).Value
Text6.Text = rs1.Fields(10).Value
Else
MsgBox "You have reached to the first record of Test Table"
End If
End Sub

Private Sub Command9_Click(Index As Integer)
Set conn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset
conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn1.Open
rs1.Open "select * from test", conn, adOpenKeyset, adLockOptimistic

rs1.MoveFirst
Text1.Text = rs1.Fields(0).Value
Text2.Text = rs1.Fields(1).Value
Text3.Text = rs1.Fields(2).Value
Text4.Text = rs1.Fields(3).Value

If rs1.Fields(4).Value = -1 Then
Option1.Value = True

Else
Option1.Value = False
End If
If rs1.Fields(5).Value = -1 Then
Option2.Value = True

Else
Option2.Value = False
End If
If rs1.Fields(6).Value = -1 Then
Option3.Value = True

Else
Option3.Value = False



End If
If rs1.Fields(7).Value = -1 Then
Option4.Value = True

Else
Option4.Value = False
End If
If rs1.Fields(8).Value = -1 Then
Option5.Value = True

Else
Option5.Value = False
End If
Text5.Text = rs1.Fields(9).Value
Text6.Text = rs1.Fields(10).Value


MsgBox "You have reached to the first record of Test Table"
End Sub

Private Sub CommandDel3_Click(Index As Integer)
Set conn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset
conn1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn1.Open
rs1.Open "select(PatientId)from test where PatientId=" & Text7.Text, conn, adOpenKeyset, adLockOptimistic
If rs1.EOF = True Then
MsgBox "Record Not found"
Else

rs1.Delete
rs1.Close
MsgBox ("Record Deleted Sucessfully")
End If

End Sub

Private Sub Form_Load()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from Pathology", conn, adOpenKeyset, adLockOptimistic
rs.MoveLast
Form4.Picture = LoadPicture("C:\Rahul Covid Project1\Resources\wp2386761.jpg")
End Sub


Private Sub MoveFirst_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from Pathology", conn, adOpenKeyset, adLockOptimistic

rs.MoveFirst
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
MsgBox "You have reached to the first record of Pathology Table"


End Sub

Private Sub MoveLast_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from Pathology", conn, adOpenKeyset, adLockOptimistic

rs.MoveLast
If Not rs.EOF Then

Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
End If
MsgBox "You have reached to the last record of Pathology Table"

End Sub

Private Sub MoveNext_Click(Index As Integer)
rs.MoveNext




If Not rs.EOF Then



Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Else
MsgBox "You have reached to the last record of Pathology Table"
End If

End Sub

Private Sub MovePrevious_Click(Index As Integer)
rs.MovePrevious
If Not rs.BOF Then

Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
Else

MsgBox "You have reached to the first record of Pathology Table"
End If

End Sub
