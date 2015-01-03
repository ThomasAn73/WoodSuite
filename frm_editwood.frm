VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_editwood 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Wood-Species editor"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   9735
   Icon            =   "frm_editwood.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Remove"
      Height          =   255
      Index           =   2
      Left            =   7800
      TabIndex        =   25
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add as new"
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   24
      Top             =   720
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Modify"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   23
      Top             =   720
      Width           =   1335
   End
   Begin VB.TextBox txt_species 
      Height          =   285
      Left            =   4920
      TabIndex        =   22
      Top             =   360
      Width           =   4215
   End
   Begin VB.ListBox lst_species 
      Height          =   1230
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   20
      Top             =   360
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   8160
      TabIndex        =   19
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   7080
      TabIndex        =   18
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   6000
      TabIndex        =   17
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   4920
      TabIndex        =   16
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   3840
      TabIndex        =   15
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   2760
      TabIndex        =   14
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   13
      Top             =   2760
      Width           =   2415
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   2566
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
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Index           =   4
      Left            =   8400
      TabIndex        =   4
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save changes"
      Height          =   375
      Index           =   3
      Left            =   7080
      TabIndex        =   3
      Top             =   5160
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove"
      Height          =   255
      Index           =   2
      Left            =   7800
      TabIndex        =   2
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add as new"
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   1
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Modify"
      Height          =   255
      Index           =   0
      Left            =   4920
      TabIndex        =   0
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Line Line4 
      X1              =   120
      X2              =   120
      Y1              =   1920
      Y2              =   3480
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   9600
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line2 
      X1              =   9600
      X2              =   9600
      Y1              =   3480
      Y2              =   1920
   End
   Begin VB.Line Line1 
      X1              =   9600
      X2              =   120
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "Wood Species"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Modulus of elasticity (E)"
      Height          =   615
      Index           =   6
      Left            =   8160
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Compression parallel to grain (Fc)"
      Height          =   735
      Index           =   5
      Left            =   7080
      TabIndex        =   11
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Compression perpendicular to grain (Fcp)"
      Height          =   735
      Index           =   4
      Left            =   6000
      TabIndex        =   10
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Shear parallel to grain (Fv)"
      Height          =   735
      Index           =   3
      Left            =   5040
      TabIndex        =   9
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Tension parallel to grain (Ft)"
      Height          =   855
      Index           =   2
      Left            =   3840
      TabIndex        =   8
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Bending (Fb)"
      Height          =   855
      Index           =   1
      Left            =   2760
      TabIndex        =   7
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Commercial Grade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   2040
      Width           =   1935
   End
End
Attribute VB_Name = "frm_editwood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim listview_selected As Boolean
Dim changed As Boolean

'INEPENDENT
'Do these steps before the form is displayed
Private Sub form_Load()
Call init_list  'Proper number of columns for the listview control
Call grade_interface(False)
Set tree = Nothing  'Clean the matrix so that it will be renewed
Call load_wood_data
changed = False
End Sub

'GRADES IMNTERFACE---------------------------------
'Editing the grades
Private Sub Command1_Click(index As Integer)
Select Case index
    Case 0  'Modify
        Call modify_grade
    Case 1  'Add-as-new
        Call add_new_grade
    Case 2  'Remove
        Call remove_grade
    Case 3  'Save
        Call save_wood_species
    Case 4  'Close
        Call prompt_to_save
        Unload frm_editwood
    End Select
End Sub
'----------------------------------------------------

'Child of grades interface
'INITIALIZE to the proper number of columns for the listview control
Private Sub init_list()
Dim columnX As ColumnHeader, run As Integer
ListView1.HideColumnHeaders = True
'first column
Set columnX = ListView1.ColumnHeaders.Add()
    columnX.Text = "Column"
    columnX.width = 2500
'the rest of the columns
For run = 2 To 7
    Set columnX = ListView1.ColumnHeaders.Add()
        columnX.Text = "Column" & run
        columnX.width = 1075
        columnX.Alignment = lvwColumnRight
Next
End Sub

'child of grades interface
'disable or enable all controls associated with Grade data
Private Sub grade_interface(status As Boolean)
Dim runnner As Integer
Command1(0).Enabled = status
Command1(1).Enabled = status
Command1(2).Enabled = status
ListView1.Enabled = status
For runner = 0 To 6
    Label2(runner).Enabled = status
    Text1(runner).Enabled = status
Next
End Sub

'when the "Modify grade" button is pressed
Private Sub modify_grade()
Dim entry As String
Dim fields(6) As String
Dim the_data As String
Dim where As String
Dim this_grade As String
Dim cur_selection As ListItem
Dim run As Integer

entry = Text1(0).Text   'the commercial grade

'is there anything selected?
If (ListView1.SelectedItem Is Nothing) Then
    mdl_beam.error (err.nonselect)
    Exit Sub
ElseIf (Trim(entry) = "") Then
    Call mdl_beam.error(err.empty_field)
    Exit Sub
End If
'therefore there is something selected and it is...
Set cur_selection = ListView1.SelectedItem

'Check to see if the proposed values are still numeric
If (numeric = False) Then Exit Sub

'prompt the user to confirm the change and then decide
If (mdl_beam.message(Update, cur_selection.Text) = Cancel) Then
    Exit Sub
End If

'check to see if the grade is still the same
'otherwise check to see if the new name already exist in the list
If (cur_selection.Text <> entry) Then
    If (mdl_beam.ItIsUnique(entry, ListView1, 0) = False) Then
        Exit Sub
    End If
End If

'We are clear to do the change
changed = True

'first remove any previous data from the matrix
'The matrix is a two dimensional collection
'in other words, it is a collection of collections
tree.delete_grade lst_species.List(lst_species.ListIndex), ListView1.SelectedItem.Text

'add the new one and update the screen
fields(0) = entry
cur_selection.Text = entry
For run = 1 To 6
    fields(run) = Text1(run).Text
    cur_selection.SubItems(run) = Text1(run).Text
Next

the_data = compose_gline(fields) 'The numerical characteristics of each grade
where = lst_species.List(lst_species.ListIndex) 'The wood type as selected
'We need to send: which_species, the grade data and the grade name
tree.add_grade where, the_data, entry

End Sub

'This subroutine is called in response to the corresponding command "add new grade"
Private Sub add_new_grade()
'check to see if "grade" is empty
'check to see if the data are numeric
If (check_valid_entries() = False) Then Exit Sub

'check to see if they already exist
If (mdl_beam.ItIsUnique(Text1(0), ListView1, 0) = False) Then
    Call mdl_beam.error(err.Exists, Text1(0))
    Exit Sub
End If

'We are clear to add the grade
changed = True
Call add_list_grade 'To add the entry to the screen and to the matrix

End Sub

Private Sub remove_grade()
Dim cur_selection As ListItem
Dim location_index As Integer

'Is there anything selected?
If (ListView1.SelectedItem Is Nothing) Then
    mdl_beam.error (err.nonselect)
    Exit Sub
Else
    Set cur_selection = ListView1.SelectedItem
    If (mdl_beam.message(sure, " = " & cur_selection.Text) = Cancel) Then Exit Sub
    
    'you are clear to remove the grade
    changed = True
    'remove from the matrix
    tree.delete_grade lst_species.List(lst_species.ListIndex), cur_selection.Text
    
    'Remove from the screen
    location_index = cur_selection.index
    ListView1.ListItems.Remove (location_index)
End If
End Sub

Private Sub save_wood_species()
Dim filetag As Integer
Dim counter As Integer
Dim runner As Integer

filetag = FreeFile  'locate a free tag for the file
Open species_file For Output As #filetag
For counter = 1 To tree.species_count
    'The first member of a grades collection is the species name
    If (tree.grade_count_num(counter) = 1) Then
        Write #filetag, lst_species.List(counter - 1) & vbCrLf
    Else
        For runner = 2 To tree.grade_count_num(counter)
            Write #filetag, tree.show_records(counter, runner) & vbCrLf
        Next
    End If
Next
Close #filetag

'Update the combo boxes
Call frmbeam.species_to_combo
frmbeam.Combo2.Clear
frmbeam.Combo2.Text = "None Selected"
End Sub

'Load data from file
Private Sub load_wood_data()
Dim filetag As Integer
Dim counter As Integer
Dim runner As Integer
Dim reclaim As String
Dim words() As String
Dim previous As String
Dim this_data As String

filetag = FreeFile  'locate a free tag for the file
Open species_file For Input As #filetag
Do While (Not EOF(filetag))
    this_data = ""
    Input #filetag, reclaim 'retreive one line
    words = mdl_beam.fragment(reclaim, Chr$(cutter) & vbCrLf, 7)
    'load species to the tree
    If (words(0) <> previous) Then
        tree.add_wood (words(0))
        lst_species.AddItem (words(0))
    End If
    If (Trim(words(1)) <> "") Then
        For runner = 1 To 7
            this_data = this_data & words(runner) & Chr$(cutter)
        Next
        tree.add_grade words(0), this_data, words(1)
    End If
    previous = words(0)
Loop
Close #filetag
End Sub

'Child of grades interface
'Highlight the text
Private Sub text1_GotFocus(index As Integer)
Select Case index
    Case 0 To 6
        Text1(index).SelStart = 0
        Text1(index).SelLength = Len(Text1(index).Text)
    End Select
End Sub

'child of grades interface
'if something was selected, tranfer it to the text fields
Private Sub listview1_itemClick(ByVal item As MSComctlLib.ListItem)
Dim run As Integer
Text1(run) = item.Text
For run = 1 To 6
    Text1(run) = item.SubItems(run)
Next
End Sub

'child of grades interface (make generic)
Private Sub clear_textfields()
Dim run As Integer
For run = 0 To 6
    Text1(run).Text = ""
Next
End Sub

'Child of grades interface (make generic?)
Private Function add_list_grade()
Dim where As String
Dim the_data As String
Dim this_grade As String
Dim fields(6) As String
Dim run As Integer
Dim itemx As ListItem
'Add the first item
Set itemx = ListView1.ListItems.Add()
    itemx.Text = Text1(0).Text
    itemx.Tag = Text1(0).Text
'add the subitems (data associated with the particular grade)
For run = 1 To 6
    itemx.SubItems(run) = Text1(run)
Next
'update the matrix
For run = 0 To 6
    fields(run) = Text1(run).Text
Next
the_data = compose_gline(fields) ' the numerical characteristics of each grade
where = txt_species.Text    'the wood type as selected
this_grade = Text1(0).Text  'the grade (name)
tree.add_grade where, the_data, this_grade

End Function

'child of grades interface  (make generic?)
'check to see if the input is acceptable
Private Function check_valid_entries() As Boolean
If (Trim(Text1(0).Text) = "") Then
    Call mdl_beam.error(err.empty_field)
    check_valid_entries = False
    Exit Function
Else
    check_valid_entries = numeric
End If
End Function

'child of "Grades interface"    (make generic!)
'check to see if the data look like numbers
Private Function numeric() As Boolean
Dim temp As Integer
Dim theValue As Single
For temp = 1 To 6
    theValue = Val(Text1(temp).Text)
    numeric = IsNumeric(Text1(temp).Text)
    If (numeric = False) Then
        Call mdl_beam.error(err.notnumber)
        Exit Function
    End If
    If (theValue <= 0) Then
        Call mdl_beam.error(err.Minusfield, Label2(temp).Caption)
        numeric = False
        Exit Function
    End If
Next
End Function

'Child of the grades interface
'add up all the text fields into one line to be stored into the matrix
Public Function compose_gline(field_data() As String, Optional start_at As Integer) As String
Dim run As Integer

For run = start_at To UBound(field_data)
    one_entry = one_entry & field_data(run) & Chr$(cutter)
Next
compose_gline = one_entry
End Function

'SPECIES INTERFACE----------------------------------
'Editing the wood types
Private Sub command2_Click(index As Integer)
Select Case index
    Case 0  'Modify
        Call modify_species
    Case 1  'Add as new
        Call add_new_species
    Case 2  'Remove
        Call remove_species
End Select
End Sub
'----------------------------------------------------

Private Sub remove_species()
Dim entry As String
Dim cur_selection As String

entry = txt_species.Text
cur_selection = lst_species.List(lst_species.ListIndex)

If (lst_species.ListIndex >= 0) Then 'if there is something selected
    If (mdl_beam.message(ask.sure2) = 1) Then
        Call grade_interface(False)
        'update the matrix
        tree.delete_species (cur_selection)
        'update the list
        lst_species.RemoveItem (lst_species.ListIndex)
        'clear the text box
        txt_species.Text = ""
        'clear the list view
        ListView1.ListItems.Clear
        changed = True
    Else
        Exit Sub
    End If
Else
    Call mdl_beam.error(err.nonselect)
End If

End Sub

Private Sub add_new_species()
Dim entry As String
entry = txt_species.Text

'check the entry to see if it is empty
If (Trim(entry) = "") Then
    Call mdl_beam.error(err.no_entry)
    Exit Sub
End If
'check if it already exists
If (already_exists2(entry) = True) Then Exit Sub

'add to the list
lst_species.AddItem entry
'add to the matrix
tree.add_wood (entry)
'automaticaly show selected item
lst_species.selected(lst_species.NewIndex) = True
changed = True
End Sub

Private Sub modify_species()
Dim entry As String
Dim cur_selection As String

entry = txt_species.Text
cur_selection = lst_species.List(lst_species.ListIndex)
'is there anything selected?
If (lst_species.ListIndex < 0) Then
    Call mdl_beam.error(err.nonselect)
    Exit Sub
ElseIf (Trim(txt_species.Text) = "") Then
    Call mdl_beam.error(err.empty_field)
    Exit Sub
End If

If (cur_selection <> entry) Then
    If (already_exists2(entry) = True) Then Exit Sub
End If

'Prompt the user for the change and decide
If (mdl_beam.message(Update, cur_selection) = Cancel) Then Exit Sub

'we are clear to do the change
lst_species.List(lst_species.ListIndex) = entry
'update the matrix
tree.update_species cur_selection, entry
changed = True

End Sub

'child of species interface (dublicate???)
'Look inside the list and see if the item already exists
Private Function already_exists2(look_up As String) As Boolean
Dim runner As Integer
For runner = 0 To lst_species.ListCount - 1
    If (lst_species.List(runner) = look_up) Then
        Call mdl_beam.error(err.Exists, "(" & txt_species.Text & ") ")
        already_exists2 = True
        Exit Function
    End If
Next
already_exists2 = False

End Function

'child of species interface
'Understand the user's selection from the list and show proper grade info
Private Sub lst_species_click()
Dim selection_is As String
If (lst_species.ListCount > 0) Then
    selection_is = lst_species.List(lst_species.ListIndex)
    txt_species.Text = selection_is
    Call grade_interface(True)
    Call mdl_beam.extract_grades(ListView1, tree, selection_is)
End If
End Sub

'Child of species interface
Private Sub txt_species_GotFocus()
txt_species.SelStart = 0
'highlight the text
txt_species.SelLength = Len(txt_species.Text)

End Sub

Public Sub prompt_to_save()
If (changed = True) Then
    If (mdl_beam.message(Save) = yes) Then
        Call save_wood_species
    End If
End If
Exit Sub
    
End Sub
