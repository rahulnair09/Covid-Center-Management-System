VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   9420
   ClientLeft      =   225
   ClientTop       =   450
   ClientWidth     =   1800
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "Form3.frx":0000
   ScaleHeight     =   10935
   ScaleWidth      =   15120
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   8400
      Top             =   7440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmd13 
      Caption         =   "DataReport"
      Height          =   375
      Left            =   15240
      TabIndex        =   42
      Top             =   9840
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Enter The Patient Id to Delete or Update the record"
      Height          =   495
      Index           =   0
      Left            =   7440
      TabIndex        =   41
      Top             =   8160
      Width           =   3255
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Database  Refresh"
      Height          =   375
      Left            =   13920
      TabIndex        =   40
      Top             =   9840
      Width           =   1215
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   11280
      TabIndex        =   39
      Top             =   8160
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      Height          =   8055
      Index           =   0
      Left            =   0
      TabIndex        =   10
      Top             =   840
      Width           =   7095
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   3360
         TabIndex        =   26
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   3360
         TabIndex        =   25
         Top             =   3720
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3360
         TabIndex        =   24
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3360
         TabIndex        =   23
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   3360
         TabIndex        =   22
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Index           =   1
         Left            =   3240
         TabIndex        =   21
         Top             =   6840
         Width           =   2655
         Begin VB.OptionButton Option4 
            Caption         =   "No"
            Height          =   195
            Left            =   1320
            TabIndex        =   34
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Yes"
            Height          =   315
            Left            =   120
            TabIndex        =   33
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Height          =   735
         Index           =   0
         Left            =   3240
         TabIndex        =   20
         Top             =   5640
         Width           =   2655
         Begin VB.OptionButton Option2 
            Caption         =   "No"
            Height          =   195
            Left            =   1320
            TabIndex        =   32
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Yes"
            Height          =   375
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1215
         Left            =   3240
         TabIndex        =   17
         Top             =   4200
         Width           =   2655
         Begin VB.CheckBox Check4 
            Caption         =   "Throat Pain"
            Height          =   375
            Left            =   1320
            TabIndex        =   30
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Body Tiredness"
            Height          =   195
            Left            =   240
            TabIndex        =   29
            Top             =   720
            Width           =   975
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Fever"
            Height          =   195
            Left            =   1440
            TabIndex        =   28
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Coughing"
            Height          =   195
            Left            =   240
            TabIndex        =   27
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Camed From Tour"
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
         TabIndex        =   19
         Top             =   7080
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Person has contacted to Red Zone"
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
         Index           =   4
         Left            =   120
         TabIndex        =   18
         Top             =   5640
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   495
         Index           =   3
         Left            =   120
         TabIndex        =   16
         Top             =   4440
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   495
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   3720
         Width           =   2655
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
         TabIndex        =   14
         Top             =   3000
         Width           =   2655
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
         Left            =   120
         TabIndex        =   13
         Top             =   2160
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
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   2415
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
         Left            =   120
         TabIndex        =   11
         Top             =   720
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   9840
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Show"
      Height          =   6015
      Left            =   7320
      TabIndex        =   6
      Top             =   720
      Width           =   7935
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   4095
         Left            =   0
         TabIndex        =   7
         Top             =   480
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   7223
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
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Menus"
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   9480
      Width           =   16815
      Begin VB.CommandButton Command8 
         Caption         =   "Move Previous"
         Height          =   375
         Index           =   3
         Left            =   12600
         TabIndex        =   38
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Move Next"
         Height          =   375
         Index           =   2
         Left            =   11160
         TabIndex        =   37
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Move Last"
         Height          =   375
         Index           =   1
         Left            =   9720
         TabIndex        =   36
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Move First"
         Height          =   375
         Index           =   0
         Left            =   8400
         TabIndex        =   35
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Back"
         Height          =   375
         Index           =   5
         Left            =   6960
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Save"
         Height          =   375
         Index           =   4
         Left            =   5640
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Clear"
         Height          =   375
         Index           =   3
         Left            =   4320
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Delete"
         Height          =   375
         Index           =   2
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command88 
         Caption         =   "Update"
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Send Details of Patient to Pathology Lab"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   8
      Top             =   120
      Width           =   8655
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Public conn As New ADODB.Connection
 Public rs As New ADODB.Recordset




Private Sub cmd13_Click()
Dim a As Integer
a = InputBox("Enter the Id")
DataEnvironment1.Command1 (a)
DataReport1.Show
DataReport1.Refresh
DataEnvironment1.rsCommand1.Close
End Sub

Private Sub Command1_Click()


Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"
conn.Open
Dim a As Integer
Dim str As String
rs.Open "select max (PatientId)from Pathology", conn, adOpenKeyset, adLockOptimistic

If IsNull(rs.Fields(0)) Then
Text1.Text = 1
Else
a = rs.Fields(0)
Text1.Text = a + 1
End If
rs.Close

End Sub

Private Sub Command10_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
Dim a As Integer
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"
conn.Open

rs.Open "select(PatientId)from Pathology where PatientId=" & Text6.Text, conn, adOpenKeyset, adLockOptimistic
If rs.EOF <> True Or rs.BOF <> True Then
MsgBox ("No:")
Exit Sub

Else
Text1.Text = rs.Fields(0)
rs.Close
End If
rs.Close


End Sub

Private Sub Command11_Click()

Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Bikedata.mdb;Persist Security Info=False"
conn.Open
rs.CursorLocation = adUseClient
rs.Open "select * from Pathology", conn, adOpenKeyset, adLockOptimistic
rs.MoveLast

Set DataGrid1.DataSource = rs
DataGrid1.Refresh
Set rs = Nothing

End Sub

Private Sub Command15_Click(Index As Integer)
MDIForm1.Hide
Form3.Hide
Form2.Show

End Sub

Private Sub Command2_Click(Index As Integer)
rs.Open "select * from Pathology where PatientId='" & Text6.Text & "'", conn, adOpenDynamic, adLockOptimistic

If rs.EOF Then
MsgBox "Record  not found:"
Else



rs.Fields(0) = Text1.Text


rs.Fields(1) = Text2.Text
rs.Fields(2) = Text3.Text
rs.Fields(3) = Text4.Text
rs.Fields(4) = Text5.Text
rs.Fields(5) = Check1.Value
rs.Fields(6) = Check2.Value
rs.Fields(7) = Check3.Value
rs.Fields(8) = Check4.Value
rs.Fields(9) = Option1.Value
rs.Fields(10) = Option2.Value
rs.Fields(11) = Option3.Value
rs.Fields(12) = Option4.Value
rs.Update
MsgBox ("Record Updated Succesfully")
rs.Close
End Sub


Private Sub Command3_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Bikedata.mdb;Persist Security Info=False"
conn.Open

rs.Open "select(PatientId)from Pathology where PatientId=" & Text6.Text, conn, adOpenKeyset, adLockOptimistic
If rs.EOF = True Then
MsgBox "Record Not found"
Else


rs.Delete
rs.Close
MsgBox ("Record Deleted Sucessfully")
End If

End Sub

Private Sub Command4_Click(Index As Integer)
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
Option1.Value = False

Option2.Value = False
Option3.Value = False
Option4.Value = False


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
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from Pathology", conn, adOpenKeyset, adLockOptimistic
rs.AddNew
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text2.Text
rs.Fields(2) = Text3.Text
rs.Fields(3) = Text4.Text
rs.Fields(4) = Text5.Text
rs.Fields(5) = Check1.Value
rs.Fields(6) = Check2.Value
rs.Fields(7) = Check3.Value
rs.Fields(8) = Check4.Value
rs.Fields(9) = Option1.Value
rs.Fields(10) = Option2.Value
rs.Fields(11) = Option3.Value
rs.Fields(12) = Option4.Value
rs.Update
MsgBox ("Record Added Succesfully"), vbExclamation


End Sub

Private Sub Command6_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from Pathology", conn, adOpenKeyset, adLockOptimistic

rs.MoveLast
If Not rs.EOF Then

Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
Text5.Text = rs.Fields(4).Value
If rs.Fields(5).Value = -1 Then
Check1.Value = 1

Else
Check1.Value = 0


End If
If rs.Fields(6).Value = -1 Then
Check2.Value = 1

Else
Check2.Value = 0
End If
If rs.Fields(7).Value = -1 Then
Check3.Value = 1

Else
Check3.Value = 0
End If
If rs.Fields(8).Value = -1 Then
Check4.Value = 1

Else
Check4.Value = 0
End If
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
End If
MsgBox " You have reached last record"

End Sub

Private Sub Command7_Click(Index As Integer)

rs.MoveNext



If Not rs.EOF Then



Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
Text5.Text = rs.Fields(4).Value
If rs.Fields(5).Value = -1 Then
Check1.Value = 1

Else
Check1.Value = 0


End If
If rs.Fields(6).Value = -1 Then
Check2.Value = 1

Else
Check2.Value = 0
End If
If rs.Fields(7).Value = -1 Then
Check3.Value = 1

Else
Check3.Value = 0
End If
If rs.Fields(8).Value = -1 Then
Check4.Value = 1

Else
Check4.Value = 0
End If
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
Else
MsgBox "No more record...Please Traverse back or else you can enter a new record"

End If
End Sub

Private Sub Command8_Click(Index As Integer)

rs.MovePrevious
If Not rs.BOF Then

Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
Text5.Text = rs.Fields(4).Value
If rs.Fields(5).Value = -1 Then
Check1.Value = 1

Else
Check1.Value = 0


End If
If rs.Fields(6).Value = -1 Then
Check2.Value = 1

Else
Check2.Value = 0
End If
If rs.Fields(7).Value = -1 Then
Check3.Value = 1

Else
Check3.Value = 0
End If
If rs.Fields(8).Value = -1 Then
Check4.Value = 1

Else
Check4.Value = 0
End If
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
Else

MsgBox "No more records.."

End If
End Sub

Private Sub Command88_Click(Index As Integer)
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
rs.Open "select * from Pathology where PatientId=" & Text6.Text, conn, adOpenKeyset, adLockOptimistic
If rs.EOF = True Then
MsgBox "Record Not found.."
rs.Fields(0).Value = Text1.Text
rs.Fields(1).Value = Text2.Text
rs.Fields(2).Value = Text3.Text
rs.Fields(3).Value = Text4.Text
rs.Fields(4).Value = Text5.Text
rs.Fields(5).Value = Check1.Value
rs.Fields(6).Value = Check2.Value
rs.Fields(7).Value = Check3.Value
rs.Fields(8).Value = Check4.Value
rs.Fields(9).Value = Option1.Value
rs.Fields(10).Value = Option2.Value
rs.Fields(11).Value = Option3.Value
rs.Fields(12).Value = Option4.Value
rs.Update
MsgBox "Record Updated Sucessfully"
rs.Close
End If

End Sub

Private Sub Command9_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Rahul Covid Project1\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from Pathology", conn, adOpenKeyset, adLockOptimistic
rs.MoveFirst
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
Text5.Text = rs.Fields(4).Value
If rs.Fields(5).Value = -1 Then
Check1.Value = 1

Else
Check1.Value = 0


End If
If rs.Fields(6).Value = -1 Then
Check2.Value = 1

Else
Check2.Value = 0
End If
If rs.Fields(7).Value = -1 Then
Check3.Value = 1

Else
Check3.Value = 0
End If
If rs.Fields(8).Value = -1 Then
Check4.Value = 1

Else
Check4.Value = 0
End If
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

MsgBox "You have reached First Record"

End Sub

Private Sub Form_Load()
Form3.Picture = LoadPicture("E:\Rahul Covid Project1\Resources\1.jpg")
End Sub
