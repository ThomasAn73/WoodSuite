VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_floatKedit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Factor Input"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lstCond 
      Height          =   2175
      Left            =   1440
      TabIndex        =   3
      Top             =   480
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3000
      MaxLength       =   5
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.PictureBox condImage 
      BackColor       =   &H00FFFFFF&
      Height          =   2535
      Left            =   120
      ScaleHeight     =   2475
      ScaleWidth      =   1155
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Effective length (Ke)"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frm_floatKedit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modeled and Rendered in "Rhino1.1"
'Just for reference: These are the precice RGB colors for the renderings
'Background=215,224,251
'Straight beam=152,152,152
'Bend beam= 0,0,140
'pin=252,255,0

Private Sub Form_Activate()
Call mdl_beam.selectText(Text1)
Call selectProper
End Sub

Private Sub Form_Deactivate()
Call Form_Unload(1)
End Sub

Private Sub Form_Load()
Top = frmbeam.Top + (frmbeam.Height - frmbeam.ScaleHeight) + frmbeam.lstFactors.Top + frmbeam.lstFactors.Height - 50
Left = frmbeam.Left + frmbeam.lstFactors.Left + 50
Call listConditions
condImage.Picture = LoadPicture(App.Path & "\condition0.jpg")

End Sub

'try to identify the text with items on the list
Private Sub selectProper()
Dim run As Integer
Dim LookingAt As String
For run = 1 To lstCond.ListItems.Count
    LookingAt = lstCond.ListItems(run).SubItems(1)
    If (Val(LookingAt) = Val(Text1.Text)) Then
        Call lstCond_ItemClick(lstCond.ListItems(run))
    End If
Next
End Sub

'Load the listView control with end conditions

Private Sub listConditions()
Dim columnX As ColumnHeader, run As Integer
Dim itemx As ListItem
lstCond.HideColumnHeaders = True

'first column
Set columnX = lstCond.ColumnHeaders.Add()
    columnX.Text = "Column"
    columnX.width = lstCond.width * 6.3 / 11
'second column
Set columnX = lstCond.ColumnHeaders.Add()
    columnX.Text = "Column"
    columnX.width = lstCond.width * 3.3 / 11
    columnX.Alignment = lvwColumnRight
    
' Enter the items in the list
Set itemx = lstCond.ListItems.Add()
    itemx.Text = "Fixed ends"
    itemx.SubItems(1) = "0.65"
Set itemx = lstCond.ListItems.Add()
    itemx.Text = "Pinned top"
    itemx.SubItems(1) = "0.80"
Set itemx = lstCond.ListItems.Add()
    itemx.Text = "Translation top"
    itemx.SubItems(1) = "1.20"
Set itemx = lstCond.ListItems.Add()
    itemx.Text = "Pinned ends"
    itemx.SubItems(1) = "1.00"
Set itemx = lstCond.ListItems.Add()
    itemx.Text = "Free top"
    itemx.SubItems(1) = "2.10"
Set itemx = lstCond.ListItems.Add()
    itemx.Text = "Translation/pinned"
    itemx.SubItems(1) = "2.00"
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload frm_floatKedit
End Sub

'When the user makes a selection
Private Sub lstCond_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim TheIndex As Integer
condImage.Cls
TheIndex = Item.Index
Select Case TheIndex
    Case 1 To 6
        condImage.Picture = LoadPicture(App.Path & "\condition" & TheIndex & ".jpg")
        Text1.Text = Item.SubItems(1)
    End Select
End Sub

Private Sub Text1_Change()
If (IsNumeric(Text1.Text) = True) Then
    frmbeam.lstFactors.SelectedItem.SubItems(2) = Text1.Text
End If

End Sub

'If the user clicks on the text, it means they want a custom value
'which does not exist in the list
'therefore, erase any pictures from the pictureBox
Private Sub Text1_Click()
Call mdl_beam.selectText(Text1)
condImage.Picture = LoadPicture(App.Path & "\condition0.jpg")

End Sub
