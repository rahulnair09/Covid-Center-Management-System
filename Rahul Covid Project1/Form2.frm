VERSION 5.00
Begin VB.Form Form2 
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   2250
   ClientWidth     =   15120
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   9045
   ScaleWidth      =   15120
   WindowState     =   2  'Maximized
   Begin VB.Menu mnupathology 
      Caption         =   "Pathology Master"
      Begin VB.Menu mnupath 
         Caption         =   "Send details to Pathology center"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnureceipt 
         Caption         =   "Issue receipt for the details"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuPl 
         Caption         =   "Send details for Plasma Donor"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuadmissionanddischarge 
      Caption         =   "Hospital Master"
      Begin VB.Menu mnuadmission 
         Caption         =   "Admission"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnutreatment 
         Caption         =   "Treatment"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnudischarge 
         Caption         =   "Discharge"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu mnuquaratine 
      Caption         =   "Home Quaratine"
      Begin VB.Menu mnuaddpatient 
         Caption         =   "Add Patient"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnudeath 
      Caption         =   " Death Certficate"
      Begin VB.Menu mnuissue 
         Caption         =   "Issue Death Certificate"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnulist 
         Caption         =   "list of D.Certificate issues"
         Shortcut        =   ^J
      End
   End
   Begin VB.Menu mnuIn 
      Caption         =   "Covid Insurance"
      Begin VB.Menu mnuInsurer 
         Caption         =   "Add a Insurer"
         Shortcut        =   ^K
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Form_Load()
Form2.Picture = LoadPicture("C:\Users\Admin\Desktop\BACKGROUND\pexels-photo-3993212.jpeg")
End Sub

Private Sub mnuaddpatient_Click()
Form11.Show
End Sub

Private Sub mnuadmission_Click()
Form8.Show
End Sub

Private Sub mnudischarge_Click()
Form10.Show
End Sub

Private Sub mnuexit_Click()
MDIForm1.Hide
Form2.Hide
Form1.Show

End Sub

Private Sub mnuInsurer_Click()
Form13.Show
End Sub

Private Sub mnuissue_Click()
Form12.Show
End Sub

Private Sub mnupath_Click()

Form3.Show
End Sub

Private Sub mnuPl_Click()
Form7.Show

End Sub

Private Sub mnureceipt_Click()
Form4.Show
End Sub

Private Sub mnusearchdonor_Click()
Form14.Show
End Sub

Private Sub mnutreatment_Click()
Form9.Show
End Sub

