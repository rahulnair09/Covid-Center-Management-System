VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   7830
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   15120
   LinkTopic       =   "Form12"
   Picture         =   "Form12.frx":0000
   ScaleHeight     =   7830
   ScaleWidth      =   15120
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text13 
      Height          =   285
      Left            =   13080
      TabIndex        =   44
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Access Admission Details"
      Height          =   855
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   6600
      Width           =   7815
      Begin VB.CommandButton Command1 
         Caption         =   "MoveFirst"
         Height          =   375
         Index           =   7
         Left            =   240
         TabIndex        =   28
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "MoveLast"
         Height          =   375
         Index           =   1
         Left            =   6000
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "MoveNext"
         Height          =   375
         Index           =   2
         Left            =   1920
         TabIndex        =   26
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "MovePrevious"
         Height          =   375
         Index           =   3
         Left            =   4080
         TabIndex        =   25
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
         Height          =   375
         Index           =   6
         Left            =   12120
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Menus"
      Height          =   855
      Index           =   0
      Left            =   120
      TabIndex        =   15
      Top             =   7680
      Width           =   17175
      Begin VB.CommandButton Command99 
         Caption         =   "DataReport"
         Height          =   435
         Left            =   16200
         TabIndex        =   53
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command10 
         Caption         =   "MovePrevious"
         Height          =   375
         Left            =   13080
         TabIndex        =   48
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command9 
         Caption         =   "MoveNext"
         Height          =   435
         Left            =   11280
         TabIndex        =   47
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Movelast"
         Height          =   375
         Left            =   14880
         TabIndex        =   46
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command5 
         Caption         =   "MoveFirst"
         Height          =   375
         Left            =   9360
         TabIndex        =   45
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command44 
         Caption         =   "Clear"
         Height          =   375
         Index           =   5
         Left            =   6240
         TabIndex        =   21
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Save"
         Height          =   375
         Index           =   4
         Left            =   4800
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command17 
         Caption         =   "Back"
         Height          =   375
         Index           =   3
         Left            =   7800
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Delete"
         Height          =   375
         Index           =   2
         Left            =   3240
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command77 
         Caption         =   "Update"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   375
         Index           =   0
         Left            =   240
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
      Width           =   15615
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   8760
         TabIndex        =   50
         Top             =   600
         Width           =   2055
         Begin VB.OptionButton Option2 
            Caption         =   "Critical"
            Height          =   255
            Left            =   1080
            TabIndex        =   52
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Normal"
            Height          =   255
            Left            =   0
            TabIndex        =   51
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox Text12 
         Height          =   405
         Left            =   13800
         TabIndex        =   43
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox Text11 
         Height          =   405
         Left            =   13680
         TabIndex        =   42
         Top             =   1560
         Width           =   1695
      End
      Begin VB.OptionButton Option4 
         Caption         =   "No"
         Height          =   195
         Left            =   14880
         TabIndex        =   41
         Top             =   960
         Width           =   735
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Yes"
         Height          =   195
         Left            =   13800
         TabIndex        =   40
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox Text10 
         Height          =   405
         Left            =   8880
         TabIndex        =   39
         Top             =   4320
         Width           =   1695
      End
      Begin VB.TextBox Text9 
         Height          =   405
         Left            =   8880
         TabIndex        =   38
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox Text8 
         Height          =   405
         Left            =   8880
         TabIndex        =   37
         Top             =   3000
         Width           =   1695
      End
      Begin VB.TextBox Text7 
         Height          =   405
         Left            =   8880
         TabIndex        =   36
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   8880
         TabIndex        =   35
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox Text5 
         Height          =   405
         Left            =   3000
         TabIndex        =   34
         Top             =   4200
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Height          =   405
         Left            =   3000
         TabIndex        =   33
         Top             =   3120
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   405
         Left            =   3000
         TabIndex        =   32
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3000
         TabIndex        =   31
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   3000
         TabIndex        =   30
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "No of Days"
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
         Left            =   6000
         TabIndex        =   29
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Left            =   240
         TabIndex        =   14
         Top             =   3120
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Reason for the Death"
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
         Left            =   240
         TabIndex        =   13
         Top             =   4080
         Width           =   2415
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
         Index           =   2
         Left            =   240
         TabIndex        =   12
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
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "DC ID"
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
         TabIndex        =   10
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Index           =   4
         Left            =   6000
         TabIndex        =   9
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Time of the Death"
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
         Left            =   5880
         TabIndex        =   8
         Top             =   4320
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Left            =   6000
         TabIndex        =   7
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Left            =   6000
         TabIndex        =   6
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Left            =   6000
         TabIndex        =   5
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Index           =   10
         Left            =   11160
         TabIndex        =   4
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Claim Amount for Insurance"
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
         TabIndex        =   3
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
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
         Index           =   13
         Left            =   11040
         TabIndex        =   2
         Top             =   3000
         Width           =   2415
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   120
      TabIndex        =   22
      Top             =   8640
      Width           =   15375
      _ExtentX        =   27120
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
   Begin VB.Label Label4 
      Caption         =   "Enter the PatientId to Update or Delete the Record"
      Height          =   615
      Left            =   9960
      TabIndex        =   49
      Top             =   6840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Issue a Death Certificate"
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
      Left            =   4080
      TabIndex        =   0
      Top             =   360
      Width           =   7215
   End
End
Attribute VB_Name = "Form12"
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

Private Sub Command1_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from Admission", conn, adOpenKeyset, adLockOptimistic
MsgBox "You have reached First record"
rs.MoveFirst
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(8).Value

Text4.Text = rs.Fields(17).Value
Set conn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset

rs1.Open "select * from Treatment", conn, adOpenKeyset, adLockOptimistic

Text6.Text = rs1.Fields(8).Value

Text7.Text = rs1.Fields(9).Value
Text8.Text = rs1.Fields(10).Value
Text9.Text = rs1.Fields(11).Value


End Sub

Private Sub Command10_Click()
rs2.MovePrevious
If Not rs2.BOF Then

Text1.Text = rs2.Fields(0).Value
Text2.Text = rs2.Fields(1).Value
Text3.Text = rs2.Fields(2).Value
Text4.Text = rs2.Fields(3).Value
Text5.Text = rs2.Fields(4).Value
If rs2.Fields(5).Value = -1 Then
Option1.Value = True
Else
Option1.Value = False

End If
If rs2.Fields(6).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False

End If

Text6.Text = rs2.Fields(7).Value
Text7.Text = rs2.Fields(8).Value
Text8.Text = rs2.Fields(9).Value
Text9.Text = rs2.Fields(10).Value
Text10.Text = rs2.Fields(11).Value
If rs2.Fields(12).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If
If rs2.Fields(13).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False

End If
Text11.Text = rs2.Fields(14).Value
Text12.Text = rs2.Fields(15).Value
Else
MsgBox "You have reached first record", vbInformation
End If

End Sub

Private Sub Command11_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.CursorLocation = adUseClient
rs.Open "select * from  DeathCertificate", conn, adOpenKeyset, adLockOptimistic
rs.MoveLast

Set DataGrid1.DataSource = rs
DataGrid1.Refresh
Set rs = Nothing
End Sub

Private Sub Command17_Click(Index As Integer)
Form12.Hide

Form2.Show

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
Text3.Text = rs.Fields(8).Value

Text4.Text = rs.Fields(17).Value
Else
MsgBox "You have reached last record"
End If
Set conn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset

rs1.Open "select * from Treatment", conn, adOpenKeyset, adLockOptimistic
If Not rs1.EOF Then
Text6.Text = rs1.Fields(8).Value

Text7.Text = rs1.Fields(9).Value
Text8.Text = rs1.Fields(10).Value
Text9.Text = rs1.Fields(11).Value
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

Text4.Text = rs.Fields(17).Value
End If

rs1.MoveNext
If Not rs1.EOF Then
Text6.Text = rs.Fields(8).Value
Text7.Text = rs1.Fields(9).Value
Text8.Text = rs1.Fields(10).Value
Text9.Text = rs1.Fields(11).Value
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

Text4.Text = rs.Fields(17).Value
End If
rs1.MovePrevious
If Not rs1.BOF Then

Text6.Text = rs1.Fields(8).Value
Text7.Text = rs1.Fields(9).Value
Text8.Text = rs1.Fields(10).Value
Text9.Text = rs1.Fields(11).Value
Else

 MsgBox "You have reached first record"
End If

End Sub

Private Sub Command44_Click(Index As Integer)
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
Text12.Text = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False


End Sub

Private Sub Command5_Click()
Set conn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset
conn2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn2.Open
rs2.Open "select * from DeathCertificate", conn, adOpenKeyset, adLockOptimistic

rs2.MoveFirst
MsgBox "You have reached first record ", vbInformation
Text1.Text = rs2.Fields(0).Value
Text2.Text = rs2.Fields(1).Value
Text3.Text = rs2.Fields(2).Value
Text4.Text = rs2.Fields(3).Value
Text5.Text = rs2.Fields(4).Value
If rs2.Fields(5).Value = -1 Then
Option1.Value = True
Else
Option1.Value = False

End If
If rs2.Fields(6).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False

End If

Text6.Text = rs2.Fields(7).Value
Text7.Text = rs2.Fields(8).Value
Text8.Text = rs2.Fields(9).Value
Text9.Text = rs2.Fields(10).Value
Text10.Text = rs2.Fields(11).Value
If rs2.Fields(12).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If
If rs2.Fields(13).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False

End If
Text11.Text = rs2.Fields(14).Value
Text12.Text = rs2.Fields(15).Value
End Sub

Private Sub Command6_Click(Index As Integer)
If Text2.Text = "" Then
        MsgBox ("Name is not entered"), vbInformation
        
        Text2.SetFocus
        Exit Sub
    End If
   
Set conn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset
conn2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn2.Open
rs2.Open "select * from DeathCertificate", conn, adOpenKeyset, adLockOptimistic
rs2.AddNew
rs2.Fields(0) = Text1.Text
rs2.Fields(1) = Text2.Text
rs2.Fields(2) = Text3.Text
rs2.Fields(3) = Text4.Text
rs2.Fields(4) = Text5.Text
rs.Fields(5) = Option1.Value
rs.Fields(6) = Option2.Value

rs2.Fields(7) = Text6.Text
rs2.Fields(8) = Text7.Text
rs2.Fields(9) = Text8.Text
rs2.Fields(10) = Text9.Text
rs2.Fields(11) = Text10.Text

rs.Fields(12) = Option3.Value
rs.Fields(13) = Option4.Value

rs2.Fields(14) = Text11.Text
rs2.Fields(15) = Text12.Text
rs2.Update
MsgBox ("Record Added Sucessfully")

End Sub

Private Sub Command7_Click(Index As Integer)

Set conn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset
conn2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn2.Open

rs2.Open "select(PatientId)from DeathCertificate where PatientId=" & Text13.Text, conn, adOpenKeyset, adLockOptimistic

rs2.Delete
rs2.Close
MsgBox ("Record Deleted Sucessfully")
End Sub

Private Sub Command77_Click(Index As Integer)
If Text2.Text = "" Then
        MsgBox ("Name is not entered"), vbInformation
        
        Text2.SetFocus
        Exit Sub
    End If
   
Set rs2 = New ADODB.Recordset
 
 rs2.Open "select * from DeathCertificate where PatientId=" & Text13.Text, conn, adOpenKeyset, adLockOptimistic
rs2.Fields(0) = Text1.Text
rs2.Fields(1) = Text2.Text
rs2.Fields(2) = Text3.Text
rs2.Fields(3) = Text4.Text
rs2.Fields(4) = Text5.Text
rs2.Fields(5) = Option1.Value
rs2.Fields(6) = Option2.Value

rs2.Fields(7) = Text6.Text
rs2.Fields(8) = Text7.Text
rs2.Fields(9) = Text8.Text
rs2.Fields(10) = Text9.Text
rs2.Fields(11) = Text10.Text

rs2.Fields(12) = Option3.Value
rs2.Fields(13) = Option4.Value

rs2.Fields(14) = Text11.Text
rs2.Fields(15) = Text12.Text
rs2.Update
MsgBox "Record Updated Sucessfully"
End Sub


Private Sub Command8_Click()
Set conn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset
conn2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn2.Open
rs2.Open "select * from DeathCertificate", conn, adOpenKeyset, adLockOptimistic
rs2.MoveLast

MsgBox "You have reached last record", vbInformation
If Not rs2.EOF Then
Text1.Text = rs2.Fields(0).Value
Text2.Text = rs2.Fields(1).Value
Text3.Text = rs2.Fields(2).Value
Text4.Text = rs2.Fields(3).Value
Text5.Text = rs2.Fields(4).Value
If rs2.Fields(5).Value = -1 Then
Option1.Value = True
Else
Option1.Value = False

End If
If rs2.Fields(6).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False

End If

Text6.Text = rs2.Fields(7).Value
Text7.Text = rs2.Fields(8).Value
Text8.Text = rs2.Fields(9).Value
Text9.Text = rs2.Fields(10).Value
Text10.Text = rs2.Fields(11).Value
If rs2.Fields(12).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If
If rs2.Fields(13).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False

End If
Text11.Text = rs2.Fields(14).Value
Text12.Text = rs2.Fields(15).Value

End If
End Sub

Private Sub Command9_Click()
rs2.MoveNext
If Not rs2.BOF Then

Text1.Text = rs2.Fields(0).Value
Text2.Text = rs2.Fields(1).Value
Text3.Text = rs2.Fields(2).Value
Text4.Text = rs2.Fields(3).Value
Text5.Text = rs2.Fields(4).Value
If rs2.Fields(5).Value = -1 Then
Option1.Value = True
Else
Option1.Value = False

End If
If rs2.Fields(6).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False

End If

Text6.Text = rs2.Fields(7).Value
Text7.Text = rs2.Fields(8).Value
Text8.Text = rs2.Fields(9).Value
Text9.Text = rs2.Fields(10).Value
Text10.Text = rs2.Fields(11).Value
If rs2.Fields(12).Value = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If
If rs2.Fields(13).Value = -1 Then
Option4.Value = True
Else
Option4.Value = False

End If
Text11.Text = rs2.Fields(14).Value
Text12.Text = rs2.Fields(15).Value
Else
MsgBox "You have reached last record", vbInformation
End If
End Sub

Private Sub Command99_Click()
Dim a As Integer
a = InputBox("Enter the Id")
DataEnvironment1.Command3 (a)
DataReport3.Show
DataReport3.Refresh
DataEnvironment1.rsCommand3.Close
End Sub

Private Sub Form_Load()
Form12.Picture = LoadPicture("C:\Rahul Covid Project1\Resources\877.jpg")
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from Admission", conn, adOpenKeyset, adLockOptimistic
rs.MoveLast
Set conn1 = New ADODB.Connection
Set rs1 = New ADODB.Recordset

rs1.Open "select * from Treatment", conn, adOpenKeyset, adLockOptimistic
rs1.MoveLast
Set conn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset

rs2.Open "select * from DeathCertificate", conn, adOpenKeyset, adLockOptimistic
rs2.MoveLast
End Sub

Private Sub Option3_Click()
Dim s As Long
Text11.Enabled = True
Dim l As Long
l = CLng(InputBox("Enter the Amount to Redeem..."))
Text11.Text = l
s = CLng(Text7.Text) + CLng(Text8.Text) + CLng(Text9.Text) - CLng(Text11.Text)
Text12.Text = CLng(s)
End Sub

Private Sub Option4_Click()
Text12.Enabled = True
Text11.Text = 0
Dim s As Long
s = CLng(Text7.Text) + CLng(Text8.Text) + CLng(Text9.Text)
Text12.Text = CLng(s)
End Sub
