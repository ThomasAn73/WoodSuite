VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmbeam 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Beam sizing"
   ClientHeight    =   7875
   ClientLeft      =   1125
   ClientTop       =   600
   ClientWidth     =   9975
   FillStyle       =   0  'Solid
   Icon            =   "Beam design.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   9975
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   36
      Text            =   "500"
      Top             =   1800
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSComctlLib.ListView lstFactors 
      Height          =   1215
      Left            =   5520
      TabIndex        =   33
      ToolTipText     =   "Double-click to change the value of each correction factor."
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2143
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lstSizes 
      Height          =   2895
      Left            =   6840
      TabIndex        =   31
      Top             =   2280
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
      Caption         =   "G"
      Height          =   255
      Index           =   5
      Left            =   1320
      TabIndex        =   30
      ToolTipText     =   "Click to evaluate AND Graph the current information."
      Top             =   3240
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Evaluate"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   29
      ToolTipText     =   "Click to process the current information without Graphing."
      Top             =   3240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Modify"
      Height          =   255
      Index           =   3
      Left            =   2640
      TabIndex        =   11
      ToolTipText     =   "Update the selected entry with the new data"
      Top             =   3240
      Width           =   855
   End
   Begin VB.ComboBox cbox_direction 
      Height          =   315
      ItemData        =   "Beam design.frx":030A
      Left            =   5520
      List            =   "Beam design.frx":031A
      Style           =   2  'Dropdown List
      TabIndex        =   10
      ToolTipText     =   "Indicate the direction that the force is pointing towards."
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txt_Blength 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2640
      MaxLength       =   7
      TabIndex        =   4
      Text            =   "0"
      ToolTipText     =   "Enter the length of the Beam."
      Top             =   1440
      Width           =   855
   End
   Begin MSComctlLib.ListView lstview 
      Height          =   1455
      Left            =   240
      TabIndex        =   28
      Top             =   3720
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txt_fdata 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   5520
      MaxLength       =   7
      TabIndex        =   17
      Text            =   "-90"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox txt_fdata 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   4560
      MaxLength       =   7
      TabIndex        =   9
      Text            =   "0"
      ToolTipText     =   "The magnitude of the force in pounds (or pounds per linear foot for uniform loads)"
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox txt_fdata 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   3600
      MaxLength       =   7
      TabIndex        =   8
      Text            =   "0"
      ToolTipText     =   "The location starting at the left of the beam (measured in feet)"
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton cmd_close 
      Caption         =   "Close"
      Height          =   375
      Left            =   8640
      TabIndex        =   16
      Top             =   7320
      Width           =   1095
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "Beam design.frx":0335
      Left            =   360
      List            =   "Beam design.frx":033F
      Style           =   2  'Dropdown List
      TabIndex        =   5
      ToolTipText     =   "Select the type of force (Load, Reaction)."
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add"
      Height          =   255
      Index           =   1
      Left            =   3600
      TabIndex        =   12
      ToolTipText     =   "Add the entry as a new item on the list"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   255
      Index           =   2
      Left            =   4560
      TabIndex        =   13
      ToolTipText     =   "Delete the selected entry"
      Top             =   3240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear list"
      Height          =   255
      Index           =   0
      Left            =   5520
      TabIndex        =   14
      ToolTipText     =   "Clear the entire list (for a new start)"
      Top             =   3240
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      Caption         =   "Wood species"
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
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Click to open the wood species editor"
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Lumber sizes"
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
      Left            =   6840
      TabIndex        =   15
      ToolTipText     =   "Click to open the board sizes editor"
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox txt_fdata 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   6
      ToolTipText     =   "Unless othrwise specified, a name is provided automaticaly"
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox txt_fdata 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   2640
      MaxLength       =   7
      TabIndex        =   7
      Text            =   "0"
      ToolTipText     =   "Enter '0' for point loads or non-zero for uniform loads"
      Top             =   2760
      Width           =   855
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "Beam design.frx":0353
      Left            =   2640
      List            =   "Beam design.frx":0355
      Sorted          =   -1  'True
      TabIndex        =   3
      Text            =   "Combo2"
      Top             =   720
      Width           =   2775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Beam design.frx":0357
      Left            =   2640
      List            =   "Beam design.frx":0359
      Sorted          =   -1  'True
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   360
      Width           =   2775
   End
   Begin VB.PictureBox Drawing1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   5  'Downward Diagonal
      Height          =   1935
      Left            =   240
      ScaleHeight     =   1875
      ScaleWidth      =   9435
      TabIndex        =   0
      ToolTipText     =   "The beam is considered lateraly braced at every reaction."
      Top             =   5280
      Width           =   9495
   End
   Begin VB.Label Label9 
      Caption         =   "Partitions"
      Height          =   255
      Left            =   3600
      TabIndex        =   38
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Accuracy"
      Height          =   255
      Left            =   240
      TabIndex        =   37
      Top             =   1800
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   "Sawn Lumber"
      Height          =   255
      Left            =   2640
      TabIndex        =   35
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lbl_Factors 
      Caption         =   "Factors/Constants"
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
      Left            =   5520
      TabIndex        =   34
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   32
      Top             =   7440
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "Type"
      Height          =   255
      Index           =   5
      Left            =   600
      TabIndex        =   27
      Top             =   2400
      Width           =   495
   End
   Begin VB.Label Label4 
      Caption         =   "Name"
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   26
      Top             =   2400
      Width           =   495
   End
   Begin VB.Line Line4 
      X1              =   240
      X2              =   240
      Y1              =   3600
      Y2              =   2280
   End
   Begin VB.Line Line3 
      X1              =   6720
      X2              =   240
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Line Line2 
      X1              =   6720
      X2              =   6720
      Y1              =   2280
      Y2              =   3600
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   6720
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Commercial Grade"
      Height          =   255
      Left            =   240
      TabIndex        =   25
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label5 
      Caption         =   "Requirement"
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
      Left            =   8280
      TabIndex        =   24
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Length of Board"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   1440
      Width           =   2295
   End
   Begin VB.Label Label6 
      Caption         =   "Feet (ft)"
      Height          =   255
      Left            =   3600
      TabIndex        =   22
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Span"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   21
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Direction"
      Height          =   255
      Index           =   4
      Left            =   5520
      TabIndex        =   20
      Top             =   2400
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "Magnitude"
      Height          =   255
      Index           =   3
      Left            =   4560
      TabIndex        =   19
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Location"
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   18
      Top             =   2400
      Width           =   855
   End
End
Attribute VB_Name = "frmbeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim feditor As New Collection
Dim group_counter As New ld_counter
Dim group_stack As New ld_stack
Enum ldTypes
    Lx = 1
    Rx = 2
    Bx = 3
End Enum

'This function is evoked every time the user makes a selection
'It is also evoked automaticaly each time we select one of its items programaticaly
Private Sub cbox_direction_Click()
Dim selection As String
Dim LoadType As String
LoadType = Combo3.Text
selection = cbox_direction.Text
Select Case selection
    Case "Up", "Down"
        'Depending on the user choices in regard to "direction", some edit fields may need to be disabled.
        If (LoadType = "Load") Then
            Call LockUnlock_LoadsInterface(True, True, True, True, True, True)
        ElseIf (LoadType = "Reaction") Then
            Call LockUnlock_LoadsInterface(True, False, True, False, False, True)
        End If
    Case "Left", "Right"
        feditor(3).Text = "0"
        If (LoadType = "Load") Then
            Call LockUnlock_LoadsInterface(True, False, True, True, True, True)
        End If
End Select
End Sub

Private Sub Combo2_Click()
Call mdl_core.determine_FsubB(Combo1.Text, Combo2.Text)
Call ReInitializeSizes
End Sub

'Do the following before the form appears on the screen
Private Sub form_Load()     'Initialise
sizes_file = App.Path & "\sizes.dat"
species_file = App.Path & "\species.dat"

'place the entry fields into a collection for easy access
feditor.Add Combo3, "1"
feditor.Add txt_fdata(0), "2"
feditor.Add txt_fdata(1), "3"
feditor.Add txt_fdata(2), "4"
feditor.Add txt_fdata(3), "5"
feditor.Add cbox_direction, "6"


Combo3.Text = "Load"
cbox_direction.Text = "Down"
factor.flat = False
factor.wet = False
factor.duration = False

Call init_forces
Call species_to_combo
Call update_grades_combo(Combo1.Text)
Call mdl_beam.charge_LumberSizes
Call init_Sizes
Call init_lstFactors
Unload frm_editsize
Call beamSchematic(factor.flat)
txt_Blength.Text = "0"
Call txt_Blength_Change

Label3.Caption = "Wood Suite (version " & App.Major & "." & App.Minor & ")"
End Sub

' Format and clean the editor fields, in a manner relevant to the choice
Private Sub prepare_fields(choice As String)
choice = Trim(choice)
' Each argument passed through the "LockUnlock_LoadsInterface" function corresponds to an entry object
Select Case choice
    Case "Reaction"
        Call LockUnlock_LoadsInterface(True, False, True, False, False, True)
        Call name_generator(Rx)
        'Customize the direction box
        cbox_direction.Text = "Up"
    Case "Load"
        Call LockUnlock_LoadsInterface(True, True, True, True, True, True)
        Call name_generator(Lx)
        'Customize the direction box
        feditor.item(6).Text = "Down"
    
    'Brace point is no longer used in the interface
    Case "Brace Point"
        Call LockUnlock_LoadsInterface(True, False, True, False, False, True)
        Call name_generator(Bx)
        feditor.item(6).Text = "Down"
        
End Select
'I suppose we could use a format with one decimal place: "0.0"
'However, it seems simpler, cleaner, with just "0"
feditor.item(3).Text = "0"
feditor.item(4).Text = "0"
feditor.item(5).Text = "0"
End Sub

' Keep track and generate names for loads/reactions/brace points
Private Sub name_generator(load_type As Integer)
Dim the_name As String
Dim previous As String
'if there is something in the stack then pick from there
'otherwise make a new one without updating the counter just yet
Select Case load_type
    Case Lx
        If (group_stack.load_count > 0) Then
            
            txt_fdata(0).Text = group_stack.load_pop
            group_stack.load_push (txt_fdata(0).Text)
        Else
            the_name = "L" & (group_counter.load_count + 1)
            txt_fdata(0).Text = the_name
            group_stack.load_push (the_name)
            
        End If
    Case Rx
        If (group_stack.reaction_count > 0) Then
            
            txt_fdata(0).Text = group_stack.reaction_pop
            group_stack.reaction_push (txt_fdata(0).Text)
        Else
            the_name = "R" & (group_counter.reaction_count + 1)
            txt_fdata(0).Text = the_name
            group_stack.reaction_push (the_name)
        End If

    Case Bx
        If (group_stack.brace_count > 0) Then
            
            txt_fdata(0).Text = group_stack.brace_pop
            group_stack.brace_push (txt_fdata(0).Text)
        Else
            the_name = "B" & (group_counter.brace_count + 1)
            txt_fdata(0).Text = the_name
            group_stack.brace_push (the_name)
        End If

End Select
End Sub

'clicking for a load type
Private Sub Combo3_Click()
Dim selection As String
selection = Combo3.List(Combo3.ListIndex)

'Enable or disable (and format) the proper text fields
Call prepare_fields(selection)

'If (lstview.ListItems.Count > 0) Then lstview.SelectedItem.selected = False
End Sub

'Data command interface
Private Sub Command1_Click(index As Integer)
Dim cur_selection As Integer
Dim EvalSuccess As Boolean
Dim assist As String

Select Case index
    Case 0:     'clear list
    ' ask for confirmation since this is a very ambitious move
        If (mdl_beam.message(ask.clear_all) = yes) Then
            Call loads_clear
            Call ReInitializeSizes
        End If
    Case 1:     'Add

        Call loads_add
        'Deselect the darn thing
        lstview.SelectedItem.selected = False
        Call refressList
        Call ReInitializeSizes
    Case 2:     'Delete
        Call loads_remove
        Call refressList
        Call ReInitializeSizes
    Case 3:     'modify
        Call loads_set
        'it seems reasonable to diselect the item which was just operated on
        'this way the interface seems more streamlined
        'it prevents the possibility of accidentaly changing the same item twice
        '(we are forcing the user to select again)
        If (lstview.ListItems.Count > 0) Then
            lstview.SelectedItem.selected = False
        End If
        Call refressList
        Call ReInitializeSizes
    Case 4, 5:    ' Evaluate, Graphs
        'Warn the user if no species or grades have been selected
        If (Left(Combo1.Text, 1) = " " Or Left(Combo2.Text, 1) = " ") Then
            Call mdl_beam.error(err.noSpecies)
            Exit Sub
        End If
        'Evaluate without graphing
        If (index = 4) Then
            Call mdl_core.Evaluate(frmbeam.lstview, Val(txt_Blength.Text), 750, False)
        
        Else                'Include graph
            Call mdl_core.Evaluate(frmbeam.lstview, Val(txt_Blength.Text), 750, True)
        End If
End Select

Call redrawForces
End Sub

'if we enter this function it means that the graphic picture may have outdated information
'Erase the old and draw the new arrangements
Private Sub redrawForces()
Dim population As Integer
Dim run As Integer
Dim thisType As String
Dim thisLocation As Single
Dim thisSpan As Single
Dim thisDirection As String
Dim thisMagnitude As String
Dim thisName As String

'make a clean start
Drawing1.Cls
Call beamSchematic(factor.flat)

population = lstview.ListItems.Count
If (population = 0) Then Exit Sub

For run = 1 To population
    thisType = lstview.ListItems(run).Text
    thisName = lstview.ListItems(run).SubItems(1)
    thisLocation = lstview.ListItems(run).SubItems(3)
    thisSpan = Val(lstview.ListItems(run).SubItems(2))
    thisDirection = lstview.ListItems(run).SubItems(5)
    thisMagnitude = lstview.ListItems(run).SubItems(4)
    'Debug.Print "Found....."; thisType, thisLocation
    Call Draw1arrow(thisType, Val(txt_Blength.Text), thisLocation, thisSpan, thisMagnitude, thisName, thisDirection)
Next
End Sub

'Accept the suggested change
Private Sub loads_set()
Dim run As Integer
Dim sel_name As String
Dim entry_name As String
Dim itemx As ListItem

' we cannot modify if there is nothing selected
If (is_selected = False) Then
    Call mdl_beam.error(err.nonselect)
    Exit Sub
End If

Set itemx = lstview.ListItems(lstview.SelectedItem.index)

sel_name = lstview.SelectedItem.SubItems(1)
entry_name = txt_fdata(0).Text
'Check to see if the name field has changed
'if it has changed look to see if there is anything else with the same name
If (entry_name <> sel_name) Then
    If (mdl_beam.ItIsUnique(entry_name, lstview, 1) = False) Then Exit Sub
End If

    ' access each object of the force editor via the collection object
    'and find where to place the "NULL" identifier
Call placeIt(itemx)

Call prepare_fields(Combo3.Text)

End Sub

'Transfer the active contents of the editor, to the list
Private Sub placeIt(TheItem As ListItem)
Dim run As Integer

TheItem.Text = feditor.item(1).Text
TheItem.Key = txt_fdata(0).Text

'The first item is the type (which is not supposed to change), therefore start from the second
For run = 2 To feditor.Count - 1
    'Fields that are disabled need to be handled differently
    If (feditor.item(run).Enabled = True) Then
        If (run > 2 And run < feditor.Count) Then
            TheItem.SubItems(run - 1) = Format(Val(feditor.item(run).Text), TheFormat)
        Else
            TheItem.SubItems(run - 1) = feditor.item(run).Text
        End If
    Else
        TheItem.SubItems(run - 1) = "Null"
        ' For every "NULL" show a zero on the coresponding field
        If (run < feditor.Count) Then feditor.item(run).Text = "0"
    End If
Next
'Update the last field: "Direction"
TheItem.SubItems(run - 1) = feditor(feditor.Count).Text
End Sub


' Clear the list
'flash the stack
'restart the counter
' in other words... clear everything, to make a new start
Private Sub loads_clear()
lstview.ListItems.Clear
group_stack.pop_all
group_counter.restart
Combo3.Text = "Load"
End Sub

'Remove button has ben pressed
Private Sub loads_remove()
Dim cur_index As Integer
' is there anything selected?
If (is_selected() = True) Then
    'push the item in the stack and decrement the counter
    Call logistics
    cur_index = lstview.SelectedItem.index
    lstview.ListItems.Remove (cur_index)
Else
    Call mdl_beam.error(err.nonselect)
End If
End Sub

Private Sub logistics()
Dim the_item As String
the_item = lstview.SelectedItem.SubItems(1)
Select Case lstview.SelectedItem.Text
    Case "Load"
        group_stack.load_push (the_item)
        group_counter.load_dec
    Case "Reaction"
        group_stack.reaction_push (the_item)
        group_counter.reaction_dec
    Case "Brace Point"
        group_stack.brace_push (the_item)
        group_counter.brace_dec
End Select
End Sub

'Browse through the list list to find whether any item is selected
Private Function is_selected() As Boolean
Dim run As Integer
' First check to see if the list contains any itmes at all
If (lstview.ListItems.Count > 0) Then
    For run = 1 To lstview.ListItems.Count
        If (lstview.ListItems(run).selected = True) Then
            is_selected = True
            Exit Function
        End If
    Next
End If
is_selected = False
End Function

' Adding a 'force' to the list
Private Sub loads_add()
Dim run As Integer
Dim itemx As ListItem

'Check for name uniqueness
If (mdl_beam.ItIsUnique(feditor(2).Text, lstview, 1) = False) Then
    mdl_beam.error err.Exists, " (" & feditor(2).Text & ") "
    Exit Sub
End If
'Check for location uniqueness
If (mdl_beam.ItIsUnique(feditor(4).Text, lstview, 3, feditor(1).Text, 0) = False) Then
    mdl_beam.error (err.locationNonUnique)
    Exit Sub
End If

'Warn the user if a load of zero magnitude is being entered
If (feditor(1).Text = "Load" And Val(feditor(5)) = 0) Then Call mdl_beam.error(err.zeroMag)

Set itemx = lstview.ListItems.Add()
Call placeIt(itemx)

Call name_update
End Sub

Private Sub name_update()
Dim cur_choice As String
cur_choice = Trim(Combo3.Text)
' The generated name is expected to be in the stack
' if we are in here it means a command has been pressed
' therefore we can finalize the name-counting process by incrementing the proper counter
' and don't forget to pop one item from the stack
Select Case cur_choice
    Case "Load"
        ' update the counter and pop the stack
        group_counter.load_inc
        group_stack.load_pop
    Case "Reaction"
        group_counter.reaction_inc
        group_stack.reaction_pop
    Case "Brace Point"
        group_counter.brace_inc
        group_stack.brace_pop
End Select
Call prepare_fields(Combo3.Text)
End Sub

Private Sub Drawing1_DblClick()
'frm_graph.show
End Sub


'This in regard to the first combo box
' if the user makes a species selection
Private Sub Combo1_Click()
Dim the_selection As String

Call ZeroThefactors
the_selection = Combo1.List(Combo1.ListIndex)
Call update_grades_combo(the_selection)
Call ReInitializeSizes

End Sub

'locate the grades for the selected species and load them
' to the second combo box
Public Sub update_grades_combo(selected As String)
Dim population As Integer, run As Integer, words() As String
Combo2.Clear

If (selected = " None selected") Then
    Combo2.Text = " Not applicable"
    Exit Sub
Else
    Combo2.Text = " None selected"
End If
If (Trim(selected) <> "") Then
    population = tree.grade_count(selected)
    If (population > 1) Then
        For run = 2 To population
            words = fragment(tree.show_specific(selected, run), Chr$(cutter))
            Combo2.AddItem words(1)
        Next
    Else
        Combo2.Text = " No grade info"
    End If

End If

End Sub

'Charge the first combo box with species data
Public Sub species_to_combo()
Dim run As Integer
Dim species_item As String
Combo1.Clear

'we need to load the form and the form will load the data into the tree matrix
load frm_editwood
If (tree.species_count > 0) Then
    Combo1.AddItem " None selected"
    Combo1.Text = " None selected"
Else
    Combo1.Text = " No species info"
End If

For run = 1 To tree.species_count
    species_item = tree.show_species(run)
    Combo1.AddItem species_item
Next
'Unload frm_editwood
End Sub

'Invoke wood species editor
Private Sub command2_Click()
load frm_editwood
frm_editwood.show
End Sub

'Invoke sizes editor
Private Sub Command3_Click()
frm_editsize.show
End Sub

'Exit if the "close" button is pressed
Private Sub cmd_close_Click()
Unload frmbeam
End
End Sub

'Initialize to the proper number of columns for the listView control
Private Sub init_forces()
Dim columnX As ColumnHeader, run As Integer
lstview.HideColumnHeaders = True

'first column
Set columnX = lstview.ColumnHeaders.Add()
    columnX.Text = "Column" & run
    columnX.width = 1270
'the rest of the columns
For run = 1 To 5
    Set columnX = lstview.ColumnHeaders.Add()
    columnX.Text = "Column" & run
    columnX.width = 960
    columnX.Alignment = lvwColumnRight
Next

End Sub

Private Sub init_lstFactors()
Dim columnX As ColumnHeader
lstFactors.HideColumnHeaders = True

Set columnX = lstFactors.ColumnHeaders.Add
    columnX.width = lstFactors.width * 5 / 11
    'This one MUST be left aligned
    columnX.Alignment = lvwColumnLeft
    
Set columnX = lstFactors.ColumnHeaders.Add
    columnX.width = lstFactors.width * 2 / 11
    columnX.Alignment = lvwColumnCenter

Set columnX = lstFactors.ColumnHeaders.Add
    columnX.width = lstFactors.width * 3 / 11
    columnX.Alignment = lvwColumnCenter

Call fill_lstfactors
End Sub

Private Sub fill_lstfactors()

lstFactors.ListItems.Add , , "Flat Use"
lstFactors.ListItems(1).SubItems(1) = "False"
lstFactors.ListItems(1).SubItems(2) = "(Varies)"

lstFactors.ListItems.Add , , "Wet service"
lstFactors.ListItems(2).SubItems(1) = "False"
lstFactors.ListItems(2).SubItems(2) = "0.85"

lstFactors.ListItems.Add , , "Load duration"
lstFactors.ListItems(3).SubItems(1) = "False"
lstFactors.ListItems(3).SubItems(2) = "1.15"

lstFactors.ListItems.Add , , "Size Factor"
lstFactors.ListItems(4).SubItems(1) = "="
lstFactors.ListItems(4).SubItems(2) = "(Varies)"

lstFactors.ListItems.Add , , "Repetitive Member"
lstFactors.ListItems(5).SubItems(1) = "="
lstFactors.ListItems(5).SubItems(2) = "1.15"


lstFactors.ListItems.Add , , "Effective length (Ke)"
lstFactors.ListItems(6).SubItems(1) = "="
lstFactors.ListItems(6).SubItems(2) = "1.00"

lstFactors.ListItems.Add , , "Buckling factor (c)"
lstFactors.ListItems(7).SubItems(1) = "="
lstFactors.ListItems(7).SubItems(2) = "0.80"

lstFactors.ListItems.Add , , "(Kce) factor"
lstFactors.ListItems(8).SubItems(1) = "="
lstFactors.ListItems(8).SubItems(2) = "0.30"

lstFactors.SelectedItem.selected = False
End Sub

Private Sub init_Sizes()
Dim columnX As ColumnHeader
lstsizes.HideColumnHeaders = True

Set columnX = lstsizes.ColumnHeaders.Add
    columnX.width = 460
    'This one MUST be left aligned
    columnX.Alignment = lvwColumnLeft
Set columnX = lstsizes.ColumnHeaders.Add
    columnX.width = 300
    columnX.Alignment = lvwColumnCenter
Set columnX = lstsizes.ColumnHeaders.Add
    columnX.width = 460
    columnX.Alignment = lvwColumnRight
Set columnX = lstsizes.ColumnHeaders.Add
    columnX.width = 1300
    columnX.Alignment = lvwColumnCenter
    
Call ListLumberSizes
End Sub

' Load 'sizes' data into the "results" window
Public Sub ListLumberSizes()

Dim TheData() As String
Dim oneLine As String
Dim c As Integer
Dim itemx As ListItem

lstsizes.ListItems.Clear
For c = 0 To UBound(LumberSizes, 2)
    If (LumberSizes(0, c) <= LumberSizes(1, c)) Then
        Set itemx = lstsizes.ListItems.Add()
        itemx.Text = LumberSizes(0, c)    'width
        itemx.SubItems(1) = "X"
        itemx.SubItems(2) = LumberSizes(1, c) 'Thickness
        itemx.SubItems(3) = "-"
    End If
Next

'deselect any default items
If (UBound(LumberSizes, 2) > 0) Then lstsizes.SelectedItem.selected = False

End Sub

'This sub determines the required number of boards
'The currspecs variable is a public "user defined" type
Public Sub ShowNumBeams()

Dim listRun As Integer
Dim arrayRun As Integer
Dim beamDim1 As Integer, beamDim2 As Integer
Dim foundit As Boolean
Dim beamThickness As Single
Dim beamWidth As Single
Dim CorFactors As Double

    For arrayRun = 0 To UBound(LumberSizes, 2)
            'update the values of the public variable "currspecs"
            ' Section Modulus
            ' S=(b*d^2) /6
            If (factor.flat = False) Then
                currspecs.SectMod = (LumberSizes(2, arrayRun) * LumberSizes(3, arrayRun) ^ 2) / 6
            Else
                currspecs.SectMod = (LumberSizes(3, arrayRun) * LumberSizes(2, arrayRun) ^ 2) / 6
            End If
            currspecs.Area = LumberSizes(2, arrayRun) * LumberSizes(3, arrayRun)
            currspecs.beamWeight = LumberSizes(6, arrayRun)
            currspecs.beamThickness = LumberSizes(2, arrayRun)
            currspecs.beamWidth = LumberSizes(3, arrayRun)
            
            'calculate the factors excluding the repetitive member factor
            CorFactors = mdl_core.correctionFactors(arrayRun, False)
            'the public function "howMany" will return the number of boards
            'HowMany draws information from the variable currSpecs
            
            'If "howmany" returns more a value higher than "two" then reculculate the correction factors including this time the repetitive member factor
            If (mdl_core.how_many(CorFactors) > 2) Then CorFactors = mdl_core.correctionFactors(arrayRun)
            'output the value to the lumber sizes list
            lstsizes.ListItems(arrayRun + 1).ListSubItems(3).Text = mdl_core.how_many(CorFactors)
    Next

End Sub


'Clean the values that have been entered to the "lumber-size requirements" list
Private Sub ReInitializeSizes(Optional WithThis As String)
Dim run As Integer

If (WithThis = "") Then WithThis = "-"
For run = 1 To lstsizes.ListItems.Count
    lstsizes.ListItems(run).SubItems(3) = WithThis
Next

End Sub

'Draw the isometric image of a beam, either flat or on-edge
'(This is just a box, an elongated parallelepiped"
Private Sub beamSchematic(orientation As Boolean)
Dim x As Integer, y As Integer, length As Single, width As Single
Dim color As RGB_Palete
Dim thiscolor As Long

color.Blue = 170
color.Green = 170
color.Red = 170
thiscolor = RGB(color.Red, color.Green, color.Blue)

x = Drawing1.ScaleWidth
y = Drawing1.ScaleHeight
length = x * 0.7
width = y * 0.3
Drawing1.Cls


Select Case orientation
    Case 1
    'draw the board flat
    
    'draw a box
    Drawing1.Line (x * 0.29 / 2 + width / 3, y * 0.9 / 2 + width / 2.5)-Step(length, width / 4), thiscolor, B
    'draw upper left horizontal edge
    Drawing1.Line (x * 0.29 / 2 + width / 3, y * 0.9 / 2 + width / 2.5)-Step(-width / 2, -width / 2), thiscolor
    'draw back left vertical corner
    Drawing1.Line Step(0, 0)-Step(0, width / 4), thiscolor
    'draw lower left horizontal edge
    Drawing1.Line Step(0, 0)-Step(width / 2, width / 2), thiscolor
    ' draw back edge
    Drawing1.Line (x * 0.29 / 2 - width / 2 + width / 3, y * 0.9 / 2 - width / 2 + width / 2.5)-Step(length, 0), thiscolor
    Drawing1.Line Step(0, 0)-Step(width / 2, width / 2), thiscolor

    Case 0
    'Draw it "On edge"
    Drawing1.Line (x * 0.3 / 2, y * 0.7 / 2)-Step(length, width), thiscolor, B
    Drawing1.Line (x * 0.3 / 2, y * 0.7 / 2)-Step(-width / 8, -width / 8), thiscolor
    Drawing1.Line Step(0, 0)-Step(0, width), thiscolor
    Drawing1.Line Step(0, 0)-Step(width / 8, width / 8), thiscolor
    Drawing1.Line (x * 0.3 / 2 - width / 8, y * 0.7 / 2 - width / 8)-Step(length, 0), thiscolor
    Drawing1.Line Step(0, 0)-Step(width / 8, width / 8), thiscolor
    End Select
    
'draw a datum line
Drawing1.Line (0, y / 2)-(x, y / 2), QBColor(9)

End Sub

'dealocate memory when unloading the form
' and unload children forms
Private Sub Form_Unload(Cancel As Integer)
Unload frm_editsize
Unload frm_editwood
Unload frm_graph
Set group_stack = Nothing
Set group_counter = Nothing
'Debug.Print "I am unloading this form"
End Sub

Private Sub Frame2_DblClick()
frm_AllFactors.show
End Sub

Private Sub Label3_Click()
frm_about.show
End Sub


' Restore default values for the correction factors and constants
'Show tooltip text: "Double click to restore default values"
Private Sub lbl_Factors_DblClick()
'...
End Sub

Private Sub lstFactors_dblClick()
Dim selection As String
Dim value As String
Dim SelIndex As Integer

SelIndex = lstFactors.SelectedItem.index
Select Case SelIndex
    Case 0
        Exit Sub
    Case 1, 4 'Invoke the SizeFactor/FlatUse editor
        frm_editsize.show
    Case 6  'Effective length factor
        load frm_floatKedit
        frm_floatKedit.Text1.Text = lstFactors.ListItems(6).SubItems(2)
        frm_floatKedit.show
    Case Else
        selection = lstFactors.SelectedItem.Text
        value = lstFactors.SelectedItem.SubItems(2)
        load frm_float
        frm_float.Label1.Caption = selection
        frm_float.Text1.Text = value
        frm_float.show
End Select
End Sub

Private Sub lstFactors_Click()
Dim whichOne As Integer

whichOne = lstFactors.SelectedItem.index

Select Case whichOne
    Case 1  'flat use
        factor.flat = Not factor.flat     'Toggle
        lstFactors.ListItems(1).SubItems(1) = factor.flat
        Call redrawForces
    Case 2  'wet service
        factor.wet = Not factor.wet       'Toggle
        lstFactors.ListItems(2).SubItems(1) = factor.wet
    Case 3  'Load duration
        factor.duration = Not factor.duration
        lstFactors.ListItems(3).SubItems(1) = factor.duration
End Select
End Sub

Private Sub lstFactors_LostFocus()
lstFactors.SelectedItem.selected = False
End Sub

Private Sub lstSizes_LostFocus()
lstsizes.SelectedItem.selected = False
End Sub

' Transfer what was selected to the text fields
Private Sub lstview_ItemClick(ByVal item As MSComctlLib.ListItem)
Dim run As Integer
Dim onechar As String
feditor.item(1).Text = item.Text
For run = 2 To 5
onechar = Left(item.SubItems(run - 1), 1)
    If (item.SubItems(run - 1) <> "Null" And onechar <> "(") Then
        feditor.item(run).Text = item.SubItems(run - 1)
    Else
        feditor.item(run).Text = "0"
    End If
Next
If (item.SubItems(5) = "Null") Then
    feditor.item(6).Text = "Down"
Else
    feditor.item(6).Text = item.SubItems(5)
End If

End Sub

Private Sub txt_Blength_LostFocus()
Dim selection As String
Dim minRequired As Single
selection = txt_Blength.Text

minRequired = Authorized()
'Check to see that the new change doesn't conflict with pre existing data in the list
If (minRequired > Val(selection)) Then
    Call mdl_beam.error(err.listSupportFailure, selection)
    txt_Blength.Text = minRequired
    txt_Blength.SetFocus
    Exit Sub
End If

'store the length
currspecs.beamLength = Val(selection)
Call refressList
Call ReInitializeSizes
Call redrawForces
End Sub

'highlight when on focus (Force editor)
Private Sub txt_fdata_GotFocus(index As Integer)
Select Case index
    Case 0 To 4
        Call highlight_TextBox(txt_fdata(index))
End Select
End Sub

Private Sub highlight_TextBox(box As TextBox)
box.SelStart = 0
box.SelLength = Len(box.Text)
End Sub

'What to do when the user inputs a length for the board
Private Sub txt_Blength_Change()
Dim selection As Single
selection = Val(txt_Blength.Text)

If (selection <= 0) Then
    Call LockUnlock_LoadsInterface(False, False, False, False, False, False)
Else
    Call prepare_fields(Combo3.Text)
End If

End Sub

' This function checks the list to see if any of the existing forces are in conflict with the specified beam length
Private Function Authorized() As Single
Dim run As Integer
Dim population As Integer
Dim currmax As Single
population = lstview.ListItems.Count

Authorized = 0
For run = 1 To population
    currmax = Val(lstview.ListItems(run).SubItems(3)) + Val(lstview.ListItems(run).SubItems(2))
    If (currmax > Authorized) Then
        Authorized = currmax
    End If
Next


End Function

'Show the image of an arrow at the specified position
Private Sub Draw1arrow(TheType As String, board_length As Single, location As Single, spanning As Single, magnitude As String, theName As String, Optional direction As String)
Dim x As Integer, y As Integer, pseudo As Integer
Dim color As Long
Dim shaft As Single
Dim point As Single
Dim vert As Single
Dim TxHeight As Single
Dim direx As Double
Dim HasMag As Boolean

'Understand the magnitude
'Remember: reaction magnitudes appear as (number), so the first character is a parenthesis
If (Val(magnitude) > 0 Or Left(magnitude, 1) = "(") Then
    HasMag = True
End If

x = Drawing1.ScaleWidth
y = Drawing1.ScaleHeight

'define the size the rectangle which would represent a uniform load
theSpan = x * 0.7
theSpan = theSpan * (spanning / board_length)

vert = 0.3
shaft = y * 0.2
point = shaft / 3
pseudo = x * 0.7
pseudo = pseudo * (location / board_length)

'Some error handling, although if we are in here, most errors have been checked
If (location > board_length) Then location = board_length
If (board_length <= 0) Then
    Call mdl_beam.error(err.proper_size)
    Exit Sub
End If

'Set proper color for the arrow
TheType = Trim(TheType)
Select Case TheType
    Case "Reaction"
        color = QBColor(4)      'Red
        'Draw a vertical line which indicates the location of the support
        Drawing1.Line (x * 0.15 + pseudo, 0)-Step(0, y), RGB(223, 211, 201)
    Case "Load"
        'Uniform load special color
    If (spanning > 0) Then
        color = QBColor(13)
        shaft = 150
    Else
        color = QBColor(2)      'Green
    End If
    Drawing1.Line (x * 0.15 + pseudo, y * 0.45)-Step(0, y * 0.1), RGB(166, 198, 164)

End Select

'set the proper direction for the arrow
 Select Case direction
    Case "Up"
        TxHeight = 0
        vert = 1 - vert
        direx = PI * 3 / 2
    Case "Down"
        TxHeight = Drawing1.TextHeight(theName)
        direx = PI / 2
        'No change
    Case "Right"
        TxHeight = Drawing1.TextHeight(theName) * 1.1
        direx = PI
        vert = 0.47
        
        'User, no longer has the ability to set direction for the reaction
        'It is set automatically
        'If (TheType = "Reaction") Then
        '    vert = 1 - vert
        '    TxHeight = -Drawing1.TextHeight(theName) / 4
        'End If
    Case "Left"
        TxHeight = Drawing1.TextHeight(theName) * 1.1
        direx = 0
        vert = 0.47
        'If (TheType = "Reaction") Then
        '    vert = 1 - vert
        '    TxHeight = -Drawing1.TextHeight(theName) / 4
        'End If

End Select

'We need to know about magnitude because, reactions should not show as an arrow unless they are not "Null"
If (HasMag = True Or TheType = "Load") Then
    'The trunk
    Drawing1.Line (x * 0.15 + pseudo, y * vert)-Step(theSpan + Cos(direx) * shaft, -Sin(direx) * shaft), color, B
    'the point
    Drawing1.Line (x * 0.15 + pseudo, y * vert)-Step(Cos(direx - PI / 6) * point / 1.4, -Sin(direx - PI / 6) * point), color
    Drawing1.Line (x * 0.15 + pseudo, y * vert)-Step(Cos(direx + PI / 6) * point / 1.4, -Sin(direx + PI / 6) * point), color
End If

'Draw the label
Drawing1.CurrentX = x * 0.15 + pseudo
Drawing1.CurrentY = y * vert - Sin(direx) * shaft - TxHeight
Drawing1.Print theName

End Sub

' Highlight textbox "Board Length"
Private Sub txt_Blength_GotFocus()
Call highlight_TextBox(txt_Blength)
End Sub

'This is purely an error checking function
'Checks whether the data just entered are valid
'txt_fdata is an array of 4 text boxes (0 is the "name"... 3 is the "Magnitude")
Private Sub txt_fdata_LostFocus(index As Integer)
Dim theValue As Single
Dim TheTextValue As String

Select Case index
    Case 1:  'Span
        theValue = txt_fdata(1).Text
        'The point of application + the span, cannot be more than the length of the beam
        If ((theValue > Val(txt_Blength.Text) - Val(txt_fdata(2).Text)) Or (theValue < 0)) Then
            
            If (theValue >= 0) Then
                Call mdl_beam.error(err.exceed_length)
                txt_fdata(1).Text = Val(txt_Blength.Text) - Val(txt_fdata(2).Text)
            Else
                'Cannot have negative span
                Call mdl_beam.error(err.MinusSpan)
                txt_fdata(1).Text = "0"
            End If
            txt_fdata(1).SetFocus
         End If
        
    Case 2:  'location
        'check the length of the span and the location of the load
        theValue = Val(txt_fdata(2).Text)
        If ((theValue > Val(txt_Blength.Text) - Val(txt_fdata(1).Text)) Or (theValue < 0)) Then
            Call mdl_beam.error(err.exceed_length)
            If (theValue >= 0) Then
                txt_fdata(2).Text = Val(txt_Blength.Text) - Val(txt_fdata(1).Text)
            Else
                txt_fdata(2).Text = "0"
            End If
            txt_fdata(2).SetFocus
         End If
    Case 3:    'Magnitude
        theValue = Val(txt_fdata(3).Text)
        If (theValue < 0) Then
            'when the magnitude is less than zero
            ' make it positive and invert the direction
            txt_fdata(3).Text = Abs(theValue)
            TheTextValue = feditor(6).Text
            Select Case TheTextValue
                Case "Up"
                    feditor(6).Text = "Down"
                Case "Down"
                    feditor(6).Text = "Up"
                Case "Right"
                    feditor(6).Text = "Left"
                Case "Left"
                    feditor(6).Text = "Right"
            End Select
        End If
End Select

End Sub

'When something "shows" as enabled it means that it has potential to
    'change either manually or automaticaly
'This function: Enables/ disables individual fields
Private Sub LockUnlock_LoadsInterface(name As Boolean, span As Boolean, location As Boolean, magnitude As Boolean, direction As Boolean, commands As Boolean)
Dim run As Integer
'the name field
Label4(0).Enabled = commands 'This field does change, but it is automatic , that's why it shows _
                as enabled for as long as the command buttons are anabled
txt_fdata(0).Enabled = name

' the span field
Label4(1).Enabled = span
txt_fdata(1).Enabled = span

' the location field
Label4(2).Enabled = location
txt_fdata(2).Enabled = location

'the magnitude fields
Label4(3).Enabled = magnitude
txt_fdata(3).Enabled = magnitude

' the direction field
cbox_direction.Enabled = direction
Label4(4).Enabled = direction
txt_fdata(4).Enabled = direction
' The command buttons and combo box
Label4(5).Enabled = commands
Combo3.Enabled = commands
For run = 0 To Command1.UBound
    Command1(run).Enabled = commands
Next

End Sub

Public Sub ListUpdate(TheType As String, TheLocation As Double, theValue As Double, axial As Boolean)
Dim listLength As Integer
Dim run As Integer
Dim oneType As ListItem
Dim oneLocation As ListSubItems
Dim thisLocation As String
Dim thisMagnitude As String

thisLocation = Format(TheLocation, TheFormat)
thisMagnitude = Format(Abs(theValue), "(" & TheFormat & ")")

listLength = lstview.ListItems.Count
'Locate the item in the list
For run = 1 To listLength
    Set oneType = lstview.ListItems(run)
    Set oneLocation = lstview.ListItems(run).ListSubItems
    If (TheType = oneType.Text And thisLocation = oneLocation(3).Text) Then
        oneLocation(4).Text = thisMagnitude
        If (theValue >= 0 And axial = False) Then
            oneLocation(5).Text = "Down"
        ElseIf (theValue < 0 And axial = False) Then
            oneLocation(5).Text = "Up"
        ElseIf (theValue >= 0 And axial = True) Then
            oneLocation(5).Text = "Right"
        ElseIf (theValue < 0 And axial = True) Then
            oneLocation(5).Text = "Left"
        End If
    End If
Next

End Sub

'We get into this function to clean values from the list, that have been inserted by the program as output results of various calculations
'Default values are being restored
Private Sub refressList()
Dim run As Integer
Dim population As Integer
Dim discription As String
Dim accessThe As ListSubItems

population = lstview.ListItems.Count
If (population = 0) Then Exit Sub
For run = 1 To population
Set accessThe = lstview.ListItems(run).ListSubItems
discription = lstview.ListItems(run).Text
    Select Case discription
        'clean the values of all reactions
        Case "Reaction"
            accessThe(2).Text = "Null"
            accessThe(4).Text = "Null"
            accessThe(5).Text = "Up"
        'The idea of brace points is no longer supported
        'Brace point characteristics have been incorporated into the "reactions"
        Case "Brace Point"
    End Select
Next

End Sub
