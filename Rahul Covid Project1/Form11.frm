VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form11 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form11"
   ClientHeight    =   8745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9270
   LinkTopic       =   "Form11"
   Picture         =   "Form11.frx":0000
   ScaleHeight     =   8745
   ScaleWidth      =   9270
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command7 
      Caption         =   "MovePrevious"
      Height          =   375
      Left            =   12240
      TabIndex        =   40
      Top             =   10080
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Movenext"
      Height          =   375
      Left            =   10680
      TabIndex        =   39
      Top             =   10080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Movelast"
      Height          =   375
      Left            =   13800
      TabIndex        =   38
      Top             =   10080
      Width           =   1215
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   11160
      TabIndex        =   36
      Top             =   8880
      Width           =   3015
   End
   Begin VB.Frame Frame3 
      Height          =   8055
      Index           =   3
      Left            =   0
      TabIndex        =   9
      Top             =   1320
      Width           =   6255
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   2760
         TabIndex        =   33
         Top             =   6120
         Width           =   2535
         Begin VB.OptionButton Option6 
            Caption         =   "No"
            Enabled         =   0   'False
            Height          =   195
            Left            =   1200
            TabIndex        =   35
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Yes"
            Enabled         =   0   'False
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   1695
         End
      End
      Begin VB.TextBox Text8 
         Height          =   405
         Left            =   2760
         TabIndex        =   32
         Top             =   7200
         Width           =   2415
      End
      Begin VB.TextBox Text6 
         Height          =   405
         Left            =   2760
         TabIndex        =   31
         Text            =   "0"
         Top             =   5400
         Width           =   2415
      End
      Begin VB.Frame Frame4 
         Height          =   855
         Left            =   2640
         TabIndex        =   28
         Top             =   4440
         Width           =   3015
         Begin VB.OptionButton Option4 
            Caption         =   "Critical"
            Height          =   195
            Left            =   1800
            TabIndex        =   30
            Top             =   360
            Width           =   1095
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Normal"
            Height          =   195
            Left            =   360
            TabIndex        =   29
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   2760
         TabIndex        =   27
         Top             =   3960
         Width           =   2175
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   2640
         TabIndex        =   24
         Top             =   3120
         Width           =   2535
         Begin VB.OptionButton Option2 
            Caption         =   "Dr.More"
            Height          =   315
            Left            =   1440
            TabIndex        =   26
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Dr.Rathod"
            Height          =   195
            Left            =   120
            TabIndex        =   25
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   2760
         TabIndex        =   23
         Top             =   2640
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   2760
         TabIndex        =   22
         Top             =   1920
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   2760
         TabIndex        =   21
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   2760
         TabIndex        =   20
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Cost"
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
         Index           =   24
         Left            =   120
         TabIndex        =   19
         Top             =   7200
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Symptoms found on Last Day"
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
         Index           =   23
         Left            =   120
         TabIndex        =   18
         Top             =   6120
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cost of Medicine"
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
         Index           =   22
         Left            =   120
         TabIndex        =   17
         Top             =   5280
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Status shared by Dr"
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
         Index           =   21
         Left            =   120
         TabIndex        =   16
         Top             =   4560
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Number of Days"
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
         Index           =   20
         Left            =   120
         TabIndex        =   15
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Id"
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
         Index           =   19
         Left            =   120
         TabIndex        =   14
         Top             =   360
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
         Index           =   18
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number"
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
         Index           =   17
         Left            =   120
         TabIndex        =   12
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date of HQ"
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
         Index           =   16
         Left            =   120
         TabIndex        =   11
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Allocated"
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
         Index           =   15
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Menus"
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   9720
      Width           =   16815
      Begin VB.CommandButton Command9 
         Caption         =   "View Database"
         Height          =   495
         Left            =   15120
         TabIndex        =   42
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "MoveFirst"
         Height          =   375
         Left            =   9480
         TabIndex        =   37
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command66 
         Caption         =   "Back"
         Height          =   375
         Index           =   5
         Left            =   7920
         TabIndex        =   8
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Save"
         Height          =   375
         Index           =   4
         Left            =   5040
         TabIndex        =   7
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Clear"
         Height          =   375
         Index           =   3
         Left            =   3480
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Delete"
         Height          =   375
         Index           =   2
         Left            =   6480
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command88 
         Caption         =   "Update"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   6855
      Left            =   6480
      TabIndex        =   1
      Top             =   1320
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   12091
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
   Begin VB.Label Label3 
      Caption         =   "Enter The PatientId to Update or Delete the Password"
      Height          =   495
      Left            =   6720
      TabIndex        =   41
      Top             =   8880
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Patient Under Home Quaratine"
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
      Index           =   0
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   7815
   End
End
Attribute VB_Name = "Form11"
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
rs.Open "select max(PatientId)from HomeQuarantine", conn, adOpenKeyset, adLockOptimistic

If IsNull(rs.Fields(0)) Then
Text1.Text = 1
Else
a = rs.Fields(0)
Text1.Text = a + 1
End If
rs.Close
End Sub

Private Sub Command2_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from HomeQuarantine", conn, adOpenKeyset, adLockOptimistic
MsgBox "You have reached First record", vbInformation
rs.MoveFirst
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
If rs.Fields(4).Value = -1 Then
Option1.Value = True
Else
Option1.Value = False

End If

If rs.Fields(5).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False

End If
Text5.Text = rs.Fields(6).Value
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
If rs.Fields(10).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False

End If
If rs.Fields(11).Value = -1 Then
Option6.Value = True
Else
Option6.Value = False

End If
Text8.Text = rs.Fields(12).Value

Text6.Text = rs.Fields(9).Value


End Sub

Private Sub Command3_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from HomeQuarantine", conn, adOpenKeyset, adLockOptimistic
MsgBox "You have reached last record", vbInformation
rs.MoveLast
MsgBox "You have reached last record", vbInformation
If Not rs.EOF Then
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
If rs.Fields(4).Value = -1 Then
Option1.Value = True
Else
Option1.Value = False

End If

If rs.Fields(5).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False

End If
Text5.Text = rs.Fields(6).Value
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

If rs.Fields(10).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False

End If
If rs.Fields(11).Value = -1 Then
Option6.Value = True
Else
Option6.Value = False

End If
Text8.Text = rs.Fields(12).Value

Text6.Text = rs.Fields(9).Value
End If

End Sub

Private Sub Command4_Click()
rs.MoveNext
If Not rs.EOF Then
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
If rs.Fields(4).Value = -1 Then
Option1.Value = True
Else
Option1.Value = False

End If

If rs.Fields(5).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False

End If
Text5.Text = rs.Fields(6).Value
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
If rs.Fields(10).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False

End If
If rs.Fields(11).Value = -1 Then
Option6.Value = True
Else
Option6.Value = False

End If
Text8.Text = rs.Fields(12).Value

Text6.Text = rs.Fields(9).Value
Else
MsgBox "You have reached  last record", vbInformation
End If
End Sub

Private Sub Command5_Click(Index As Integer)
 Dim num As String
num = Text3.Text
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
Text3.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
        MsgBox ("Name is not entered")
        
        Text2.SetFocus
        Exit Sub
    End If
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from HomeQuarantine", conn, adOpenKeyset, adLockOptimistic
rs.AddNew
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text2.Text
rs.Fields(2) = Text3.Text
rs.Fields(3) = Text4.Text
rs.Fields(4) = Option1.Value
rs.Fields(5) = Option2.Value
rs.Fields(6) = Text5.Text
rs.Fields(7) = Option3.Value
rs.Fields(8) = Option4.Value

rs.Fields(9) = Text6.Text
rs.Fields(10) = Option5.Value
rs.Fields(11) = Option6.Value
rs.Fields(12) = Text8.Text
rs.Update
MsgBox ("Record Added Successfully")

End Sub

Private Sub Command6_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open

rs.Open "select(PatientId)from HomeQuarantine where PatientId=" & Text7.Text, conn, adOpenKeyset, adLockOptimistic

If rs.EOF = True Then
MsgBox "No Records Found"
Else

rs.Delete
rs.Close
MsgBox ("Record Deleted Sucessfully")
End If
End Sub

Private Sub Command66_Click(Index As Integer)
Form2.Show
End Sub

Private Sub Command7_Click()
rs.MovePrevious
If Not rs.BOF Then
Text1.Text = rs.Fields(0).Value
Text2.Text = rs.Fields(1).Value
Text3.Text = rs.Fields(2).Value
Text4.Text = rs.Fields(3).Value
If rs.Fields(4).Value = -1 Then
Option1.Value = True
Else
Option1.Value = False

End If

If rs.Fields(5).Value = -1 Then
Option2.Value = True
Else
Option2.Value = False

End If
Text5.Text = rs.Fields(6).Value
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
If rs.Fields(10).Value = -1 Then
Option5.Value = True
Else
Option5.Value = False

End If
If rs.Fields(11).Value = -1 Then
Option6.Value = True
Else
Option6.Value = False

End If
Text8.Text = rs.Fields(12).Value

Text6.Text = rs.Fields(9).Value
Else
MsgBox "You have reached last record", vbInformation
End If

End Sub

Private Sub Command8_Click(Index As Integer)
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False
Option4.Value = False
Option5.Value = False
Text5.Text = ""
Text6.Text = ""
Text8.Text = ""

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
Set rs = New ADODB.Recordset
 
 rs.Open "select * from HomeQuarantine where PatientId=" & Text7.Text, conn, adOpenKeyset, adLockOptimistic
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text2.Text
rs.Fields(2) = Text3.Text
rs.Fields(3) = Text4.Text
rs.Fields(4) = Option1.Value
rs.Fields(5) = Option2.Value
rs.Fields(6) = Text5.Text
rs.Fields(7) = Option3.Value
rs.Fields(8) = Option4.Value

rs.Fields(9) = Text6.Text
rs.Fields(10) = Option5.Value
rs.Fields(11) = Option6.Value
rs.Fields(12) = Text8.Text
rs.Update
MsgBox "Record Updated Succesfully"
rs.Close

End Sub


Private Sub Command9_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.CursorLocation = adUseClient
rs.Open "select * from HomeQuarantine", conn, adOpenKeyset, adLockOptimistic
rs.MoveLast

Set DataGrid1.DataSource = rs
DataGrid1.Refresh
Set rs = Nothing
End Sub

Private Sub Form_Load()
Form11.Picture = LoadPicture("C:\Rahul Covid Project1\Resources\789.jpg")
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.CursorLocation = adUseClient
rs.Open "select * from HomeQuarantine", conn, adOpenKeyset, adLockOptimistic
rs.MoveLast
End Sub

Private Sub Text5_Change()
Dim s As Integer
s = 1
Dim c As Integer
c = Text5.Text
If c = 14 Then
Option5.Enabled = True
Option6.Enabled = True
End If
For s = 1 To c
Dim t As Integer
t = t + 60
Next
Text8.Text = t + Text6.Text


End Sub

