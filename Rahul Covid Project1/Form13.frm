VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form13 
   Caption         =   "Form13"
   ClientHeight    =   8130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12435
   LinkTopic       =   "Form13"
   Picture         =   "Form13.frx":0000
   ScaleHeight     =   8130
   ScaleWidth      =   12435
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   13200
      TabIndex        =   32
      Top             =   4920
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
      Caption         =   "MovePrevious"
      Height          =   375
      Left            =   13680
      TabIndex        =   30
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Movenext"
      Height          =   375
      Left            =   11880
      TabIndex        =   29
      Top             =   6360
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Movelast"
      Height          =   315
      Left            =   15480
      TabIndex        =   28
      Top             =   6360
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MoveFirst"
      Height          =   375
      Left            =   10080
      TabIndex        =   27
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   -360
      TabIndex        =   8
      Top             =   6000
      Width           =   19335
      Begin VB.CommandButton Command7 
         Caption         =   "View Database"
         Height          =   375
         Left            =   17640
         TabIndex        =   33
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command16 
         Caption         =   "Update"
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command89 
         Caption         =   "Delete"
         Height          =   375
         Index           =   2
         Left            =   3720
         TabIndex        =   12
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Clear"
         Height          =   375
         Index           =   3
         Left            =   8880
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save"
         Height          =   375
         Index           =   4
         Left            =   5520
         TabIndex        =   10
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command88 
         Caption         =   "Back"
         Height          =   375
         Index           =   5
         Left            =   7080
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
      Height          =   5055
      Index           =   3
      Left            =   2400
      TabIndex        =   1
      Top             =   720
      Width           =   6255
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   2880
         TabIndex        =   26
         Top             =   4440
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   2880
         TabIndex        =   25
         Top             =   3840
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2880
         TabIndex        =   24
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2880
         TabIndex        =   23
         Top             =   1800
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2880
         TabIndex        =   22
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2880
         TabIndex        =   21
         Top             =   480
         Width           =   1815
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   2880
         TabIndex        =   16
         Top             =   2160
         Width           =   2775
         Begin VB.OptionButton Option3 
            Caption         =   "12MPlan"
            Height          =   195
            Left            =   1080
            TabIndex        =   19
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton Option2 
            Caption         =   "6M Plan"
            Height          =   195
            Left            =   1200
            TabIndex        =   18
            Top             =   360
            Width           =   1335
         End
         Begin VB.OptionButton Option1 
            Caption         =   "3 M Plan"
            Height          =   615
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Activation"
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
         Left            =   120
         TabIndex        =   15
         Top             =   3840
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Charges"
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
         TabIndex        =   7
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Plan "
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
         TabIndex        =   6
         Top             =   2520
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
         TabIndex        =   5
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
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
         TabIndex        =   4
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Id"
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
         TabIndex        =   3
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Active Date"
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
         TabIndex        =   2
         Top             =   4440
         Width           =   2415
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   0
      TabIndex        =   14
      Top             =   7320
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
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enter the PatientId to Update or Delete the Password"
      Height          =   615
      Left            =   8880
      TabIndex        =   31
      Top             =   4920
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Add a Covid Insurer"
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
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   9015
   End
End
Attribute VB_Name = "Form13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public rs As New ADODB.Recordset
 Public conn As New ADODB.Connection
 
Private Sub Command1_Click(Index As Integer)
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
Dim a As Integer
Dim str As String
rs.Open "select max(CustomerId)from CovidInsurance", conn, adOpenKeyset, adLockOptimistic

If IsNull(rs.Fields(0)) Then
Text1.Text = 1
Else
a = rs.Fields(0)
Text1.Text = a + 1
End If
rs.Close
End Sub

Private Sub Command16_Click(Index As Integer)
If Text2.Text = "" Then
        MsgBox ("Name is not entered"), vbInformation
        
        Text2.SetFocus
        Exit Sub
    End If
Set rs = New ADODB.Recordset
 
 
 rs.Open "select * from CovidInsurance where CustomerId=" & Text7.Text, conn, adOpenKeyset, adLockOptimistic
 rs.Fields(0) = Text1.Text
rs.Fields(1) = Text2.Text
rs.Fields(2) = Text3.Text
rs.Fields(3) = Option1.Value
rs.Fields(4) = Option2.Value
rs.Fields(5) = Option3.Value
rs.Fields(6) = Text4.Text
rs.Fields(7) = Text5.Text
rs.Fields(8) = Text6.Text
rs.Update
MsgBox "Record Updated Succesfully"
rs.Close
End Sub

Private Sub Command2_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from CovidInsurance", conn, adOpenKeyset, adLockOptimistic
MsgBox "You have reached First record", vbInformation
rs.MoveFirst
Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
If rs.Fields(3) = -1 Then
Option1.Value = True
Else
Option1.Value = False
End If
If rs.Fields(4) = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
If rs.Fields(5) = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If
Text4.Text = rs.Fields(6).Value
Text5.Text = rs.Fields(7).Value
Text6.Text = rs.Fields(8).Value

End Sub

Private Sub Command3_Click(Index As Integer)
If Text2.Text = "" Then
        MsgBox ("Name is not entered"), vbInformation
        
        Text2.SetFocus
        Exit Sub
    End If
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from CovidInsurance", conn, adOpenKeyset, adLockOptimistic
rs.AddNew
rs.Fields(0) = Text1.Text
rs.Fields(1) = Text2.Text
rs.Fields(2) = Text3.Text
rs.Fields(3) = Option1.Value
rs.Fields(4) = Option2.Value
rs.Fields(5) = Option3.Value
rs.Fields(6) = Text4.Text
rs.Fields(7) = Text5.Text
rs.Fields(8) = Text6.Text
rs.Update
MsgBox ("Record Added Sucessfully")
rs.Close
End Sub

Private Sub Command4_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from CovidInsurance", conn, adOpenKeyset, adLockOptimistic

rs.MoveLast
If Not rs.EOF Then

Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
If rs.Fields(3) = -1 Then
Option1.Value = True
Else
Option1.Value = False
End If
If rs.Fields(4) = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
If rs.Fields(5) = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If
Text4.Text = rs.Fields(6).Value
Text5.Text = rs.Fields(7).Value
Text6.Text = rs.Fields(8).Value
End If
End Sub


Private Sub Command5_Click()
rs.MoveNext
If Not rs.EOF Then

Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
If rs.Fields(3) = -1 Then
Option1.Value = True
Else
Option1.Value = False
End If
If rs.Fields(4) = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
If rs.Fields(5) = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If
Text4.Text = rs.Fields(6).Value
Text5.Text = rs.Fields(7).Value
Text6.Text = rs.Fields(8).Value
Else
MsgBox "You have reached last record", vbInformation
End If
End Sub

Private Sub Command6_Click()
rs.MovePrevious
If Not rs.BOF Then

Text1.Text = rs.Fields(0)
Text2.Text = rs.Fields(1)
Text3.Text = rs.Fields(2)
If rs.Fields(3) = -1 Then
Option1.Value = True
Else
Option1.Value = False
End If
If rs.Fields(4) = -1 Then
Option2.Value = True
Else
Option2.Value = False
End If
If rs.Fields(5) = -1 Then
Option3.Value = True
Else
Option3.Value = False
End If
Text4.Text = rs.Fields(6).Value
Text5.Text = rs.Fields(7).Value
Text6.Text = rs.Fields(8).Value
Else
MsgBox "You have reached First record", vbInformation
End If
End Sub

Private Sub Command7_Click()
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.CursorLocation = adUseClient
rs.Open "select * from CovidInsurance", conn, adOpenKeyset, adLockOptimistic
rs.MoveLast

Set DataGrid1.DataSource = rs
DataGrid1.Refresh
Set rs = Nothing
End Sub

Private Sub Command8_Click(Index As Integer)
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Option1.Value = False
Option2.Value = False
Option3.Value = False
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""

End Sub

Private Sub Command88_Click(Index As Integer)
Form13.Hide

Form2.Show


End Sub

Private Sub Command89_Click(Index As Integer)
Set conn2 = New ADODB.Connection
Set rs2 = New ADODB.Recordset
conn2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn2.Open

rs2.Open "select(CustomerId)from  where CovidInsurance=" & Text7.Text, conn, adOpenKeyset, adLockOptimistic

rs2.Delete
rs2.Close
MsgBox ("Record Deleted Sucessfully")
End Sub

Private Sub Form_Load()
Form13.Picture = LoadPicture("C:\Rahul Covid Project1\Resources\ttt.jpg")
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset
conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Admin\Documents\Database1.mdb;Persist Security Info=False"
conn.Open
rs.Open "select * from CovidInsurance", conn, adOpenKeyset, adLockOptimistic
rs.MoveLast
End Sub

Private Sub Option1_Click()
Text4.Enabled = True
Text4.Text = 2000

End Sub

Private Sub Option2_Click()
Text4.Enabled = True
Text4.Text = 4000

End Sub

Private Sub Option3_Click()
Text4.Enabled = True
Text4.Text = 6000
End Sub
