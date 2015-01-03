VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "matrix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private wood_type As New Collection
Private grades As Collection
'works
Public Sub add_wood(species As String)
Set grades = New Collection
grades.Add species
'The first item of "grades" is the name of the species it belongs to
wood_type.Add grades, species
End Sub
Public Function show_species(sp_index As Integer) As String
show_species = wood_type.Item(sp_index).Item(1)
End Function
'works
Public Function species_count() As Integer
species_count = wood_type.Count
End Function
'Remove an entire wood species with all the child information
Public Sub delete_species(which_one As String)
wood_type.Remove (which_one)
End Sub

'Browse numericaly through the matrix
Public Function show_records(num_species As Integer, num_grade As Integer) As String
show_records = wood_type.Item(num_species).Item(num_grade)
End Function
' Knowing the species browse through its grades
Public Function show_specific(species As String, num_grade As Integer) As String
show_specific = wood_type.Item(species).Item(num_grade)
End Function

'where: is for species
Public Sub add_grade(which_species As String, the_data As String, this_grade As String)
wood_type.Item(which_species).Add which_species & Chr$(182) & the_data, this_grade

End Sub
'works
Public Function grade_count(which_species As String) As Integer
grade_count = wood_type.Item(which_species).Count
End Function
Public Function grade_count_num(which_species As Integer) As Integer
grade_count_num = wood_type.Item(which_species).Count
End Function

Public Sub delete_grade(where As String, this_grade As String)
wood_type.Item(where).Remove (this_grade)
End Sub

Public Sub update_species(which_species As String, to_what As String)
Dim hold_grade As Collection
Set hold_grade = New Collection
Dim run As Integer
'Update the first element and tail the rest
hold_grade.Add to_what
For run = 2 To wood_type.Item(which_species).Count
    hold_grade.Add wood_type.Item(which_species).Item(run)
Next
'Remove the old grade
'(I wish Collections had an edit method so I wouldn't have to do this)
wood_type.Remove (which_species)
'Add it again with the new name
wood_type.Add hold_grade, to_what
End Sub

