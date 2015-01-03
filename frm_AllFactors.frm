VERSION 5.00
Begin VB.Form frm_about 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Version Info"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   5190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3840
      Width           =   4935
   End
   Begin VB.PictureBox Picture1 
      Height          =   3615
      Left            =   120
      Picture         =   "frm_AllFactors.frx":0000
      ScaleHeight     =   3555
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frm_about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
'Debug.Print "Hey"

End Sub

Private Sub Form_Deactivate()
Call Form_Unload(1)
End Sub

Private Sub form_Load()
Top = frmbeam.Top + (frmbeam.Height - frm_about.Height) / 2
Left = frmbeam.Left + (frmbeam.width - frm_about.width) / 4
Text1.Text = "Wood Suite (Version " & App.Major & "." & App.Minor & ") - Release date May/2000" & vbCrLf & _
            "Design and programming: Thomas Doehtal Anagnostou"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frm_about
End Sub

