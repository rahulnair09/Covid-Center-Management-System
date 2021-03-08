VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15120
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   MousePointer    =   15  'Size All
   Picture         =   "Form7.frx":0000
   ScaleHeight     =   8520
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      Caption         =   "MovePrev"
      Height          =   375
      Left            =   11160
      TabIndex        =   35
      Top             =   9840
      Width           =   1095
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Movelast"
      Height          =   375
      Left            =   12480
      TabIndex        =   34
      Top             =   9840
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "MoveNext"
      Height          =   375
      Left            =   9960
      TabIndex        =   33
      Top             =   9840
      Width           =   975
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   11520
      TabIndex        =   31
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Menus"
      Height          =   855
      Index           =   2
      Left            =   0
      TabIndex        =   11
      Top             =   9480
      Width           =   15375
      Begin VB.CommandButton Command9 
         Caption         =   "ViewDatabase"
         Height          =   375
         Left            =   13920
         TabIndex        =   36
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "MoveFirst"
         Height          =   375
         Left            =   8760
         TabIndex        =   32
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   375
         Index           =   17
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command77 
         Caption         =   "Update"
         Height          =   375
         Index           =   16
         Left            =   1680
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Delete"
         Height          =   375
         Index           =   15
         Left            =   3120
         TabIndex        =   15
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         Height          =   375
         Index           =   14
         Left            =   4560
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Save"
         Height          =   375
         Index           =   13
         Left            =   5880
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Back"
         Height          =   375
         Index           =   12
         Left            =   7320
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   8295
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   6255
      Begin VB.OptionButton Option4 
         Caption         =   "No"
         Height          =   255
         Left            =   4560
         TabIndex        =   30
         Top             =   7800
         Width           =   1455
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Yes"
         Height          =   495
         Left            =   3120
         TabIndex        =   29
         Top             =   7680
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   3120
         TabIndex        =   28
         Top             =   6840
         Width           =   2655
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   3240
         TabIndex        =   27
         Top             =   5880
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   3120
         TabIndex        =   24
         Top             =   4440
         Width           =   2655
         Begin VB.OptionButton Option2 
            Caption         =   "No"
            Height          =   255
            Left            =   1440
            TabIndex        =   26
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Yes"
            Height          =   195
            Left            =   240
            TabIndex        =   25
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   3240
         TabIndex        =   23
         Top             =   3840
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3240
         TabIndex        =   22
         Top             =   3120
         Width           =   2055
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3240
         TabIndex        =   21
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   3240
         TabIndex        =   20
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3240
         TabIndex        =   19
         Top             =   720
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Whether the last two reports where found negative"
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
         Index           =   10
         Left            =   120
         TabIndex        =   10
         Top             =   7440
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
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
         Left            =   120
         TabIndex        =   9
         Top             =   5880
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Recovery"
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
         TabIndex        =   8
         Top             =   6840
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Whether the Patient has recovered from our Covid Center"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Index           =   7
         Left            =   120
         TabIndex        =   7
         Top             =   4560
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
         Height          =   615
         Index           =   6
         Left            =   120
         TabIndex        =   6
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
         Index           =   5
         Left            =   120
         TabIndex        =   5
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   495
         Index           =   4
         Left            =   120
         TabIndex        =   4
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
         Index           =   3
         Left            =   120
         TabIndex        =   3
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
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   2415
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   5535
      Left            =   6360
      TabIndex        =   18
      Top             =   1320
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   9763
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
   Begin VB.Image Image2 
      Height          =   795
      Left            =   11400
      Picture         =   "Form7.frx":18142
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   795
      Left            =   2520
      Picture         =   "Form7.frx":1E755
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the PatientID to Update or Delete the Record"
      Height          =   495
      Left            =   7440
      TabIndex        =   37
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "    Send Details for the Plasma Donor to Pathology Lab"
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
      Index           =   0
      Left            =   2520
      TabIndex        =   0
      Top             =   120
      Width           =   10215
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
 Public conn As New ADODB.Connection
 Public rs As New ADODB.Recordset

Private Sub Command1_Click(Index As Integer)

Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
Dim a As Integer
Dim str As String
rs.Open "select max(PatientId)from PlasmaDonor", conn, adOpenKeyset, adLockOptimistic

If IsNull(rs.Fields(0)) Then
Text1.Text = 1
Else
a = rs.Fields(0)
Text1.Text = a + 1
End If
rs.Close
End Sub

Private Sub Command15_Click(Index As Integer)

Unload Me


MDIForm1.Hide
Form2.Show
End Sub

Private Sub Command2_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from PlasmaDonor", conn, adOpenKeyset, adLockOptimistic

rs.MoveFirst
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
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
Text6.Text = rs.Fields(7)
Text7.Text = rs.Fields(8)
If rs.Fields(9).Value = -1 Then
Option3.Value = True

Else
Option3.Value = False

End If
If rs.Fields(10).Value = -1 Then
Option4.Value = True

Else
Option4.Value = False
MsgBox "You have reached to first record"
End If
End Sub

Private Sub Command3_Click(Index As Integer)
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Text6.Text = ""
Text7.Text = ""

End Sub

Private Sub Command4_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open

rs.Open "select(PatientId)from PlasmaDonor where PatientId=" & Text8.Text, conn, adOpenKeyset, adLockOptimistic


If rs.EOF Then
MsgBox "No records Found.."
Else
rs.Delete
rs.Close
MsgBox ("Record Deleted Sucessfully")
End If
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
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from PlasmaDonor", conn, adOpenKeyset, adLockOptimistic
rs.AddNew
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text2.Text
rs.Fields(2) = Text3.Text
rs.Fields(3) = Text4.Text
rs.Fields(4) = Text5.Text
rs.Fields(5) = Option1.Value
rs.Fields(6) = Option2.Value
rs.Fields(7) = Text6.Text
rs.Fields(8) = Text7.Text
rs.Fields(9) = Option3.Value
rs.Fields(10) = Option4.Value
MsgBox ("Record Added Succesfully")
rs.Update
rs.Close

End Sub

Private Sub Command6_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from PlasmaDonor", conn, adOpenKeyset, adLockOptimistic

rs.MoveLast
If Not rs.EOF Then
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
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
Text6.Text = rs.Fields(7)
Text7.Text = rs.Fields(8)
If rs.Fields(9).Value = -1 Then
Option3.Value = True

Else
Option3.Value = False

End If
If rs.Fields(10).Value = -1 Then
Option4.Value = True

Else
Option4.Value = False
End If
Else
MsgBox "You have reached last record "

End If


End Sub

Private Sub VScroll_Change()



End Sub

Private Sub Command7_Click()
rs.MoveNext



If Not rs.EOF Then
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
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
Text6.Text = rs.Fields(7)
Text7.Text = rs.Fields(8)
If rs.Fields(9).Value = -1 Then
Option3.Value = True

Else
Option3.Value = False

End If
If rs.Fields(10).Value = -1 Then
Option4.Value = True

Else
Option4.Value = False
End If
Else
MsgBox "You have reached last record......#NomoreRecords "


End If

End Sub

Private Sub Command77_Click(Index As Integer)
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
'Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
'conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
'conn.Open
rs.Open "select * from PlasmaDonor where PatientId=" & Text8.Text, conn, adOpenKeyset, adLockOptimistic
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text2.Text
rs.Fields(2) = Text3.Text
rs.Fields(3) = Text4.Text
rs.Fields(4) = Text5.Text
rs.Fields(5) = Option1.Value
rs.Fields(6) = Option2.Value
rs.Fields(7) = Text6.Text
rs.Fields(8) = Text7.Text
rs.Fields(9) = Option3.Value
rs.Fields(10) = Option4.Value
rs.Update
rs.Close
MsgBox "Record Updated Sucessfully"

End Sub

Private Sub Command8_Click()
rs.MovePrevious
If Not rs.BOF Then
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
Text4.Text = rs.Fields(3)
Text5.Text = rs.Fields(4)
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
Text6.Text = rs.Fields(7)
Text7.Text = rs.Fields(8)
If rs.Fields(9).Value = -1 Then
Option3.Value = True

Else
Option3.Value = False

End If
If rs.Fields(10).Value = -1 Then
Option4.Value = True

Else
Option4.Value = False
End If
Else
MsgBox "You have reached first record "

End If

End Sub

Private Sub Command9_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.CursorLocation = adUseClient
rs.Open "select * from PlasmaDonor", conn, adOpenKeyset, adLockOptimistic
rs.MoveLast

Set DataGrid1.DataSource = rs
DataGrid1.Refresh
Set rs = Nothing
End Sub

Private Sub Form_Load()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from PlasmaDonor", conn, adOpenKeyset, adLockOptimistic

Form7.Picture = LoadPicture("")

End Sub

