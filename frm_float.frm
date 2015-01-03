VERSION 5.00
Begin VB.Form frm_float 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Factor Input"
   ClientHeight    =   510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3795
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   510
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2280
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frm_float"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
Text1.SetFocus
End Sub

Private Sub Form_Deactivate()
Call Form_Unload(1)
End Sub

Private Sub Form_Load()
Top = frmbeam.Top + (frmbeam.Height - frmbeam.ScaleHeight) + frmbeam.lstFactors.Top + frmbeam.lstFactors.Height - 50
Left = frmbeam.Left + frmbeam.lstFactors.Left + 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frm_float
End Sub

Private Sub Text1_Change()
If (IsNumeric(Text1.Text) = True) Then
    frmbeam.lstFactors.SelectedItem.SubItems(2) = Text1.Text
End If
End Sub
