Attribute VB_Name = "mdl_beam"

'This is an INDEPENDENT stand alone function.........
'It does the same thing as the SPLIT function (when leading and trailing spaces have benn removed with TRIM)
'---------------------------------------------------
'Receives a string along with user defined delimiters
'It brakes the line of text contained in the string into
'words (fragments) based on the given delimiters
Public Function fragment(thestring As String, Optional delimiters As String, Optional default As Integer) As String()
Dim fields() As String
Dim delim As String
Dim strlength As Integer
Dim runner As Integer, delimrunner As Integer
Dim currentchar As String, word As Integer
Dim flag As Integer 'remember the previous delimiter position
Dim found As Boolean
If (delimiters = "") Then delimiters = " "
strlength = Len(thestring)
word = -1
flag = -1

For runner = 1 To strlength 'traverse the string
    found = False
    currentchar = Mid(thestring, runner, 1)
    For delimrunner = 1 To Len(delimiters) 'Traverse the delimeters
        delim = Mid(delimiters, delimrunner, 1)
        If (delim = currentchar) Then
            found = True
        End If
    Next
If (found = False) Then
    If (runner > (flag + 1)) Then
        word = word + 1
        ReDim Preserve fields(word)
    End If
    fields(word) = fields(word) + currentchar
    flag = runner
End If
Next

If (word > 0) Then
    fragment = fields   'return results in array format
Else                    'Return a default array if there was nothing to fragment
    ReDim Preserve fields(default)
    fragment = fields
End If
End Function
' Child of frm_editsize
'Retreive board-size data from file and place them in an array for random access
Public Sub charge_LumberSizes()

Dim fields() As String
Dim reclaim As String
Dim counter As Integer, c As Integer, this_file As Integer

'default size
ReDim LumberSizes(6, 0)

counter = 0
this_file = FreeFile
Open sizes_file For Input As #this_file
Do While (Not EOF(this_file))
    Input #this_file, reclaim
    fields = fragment(reclaim, delims & Chr$(cutter), 6)
    ReDim Preserve LumberSizes(6, counter)
    ' Two dimensional array....
    'First dimension accesses the fields
    'Second dimension accesses the beam sizes
    For c = 0 To UBound(fields) 'Populate the columns
        LumberSizes(c, counter) = fields(c)
    Next
    counter = counter + 1
Loop
Close #this_file
End Sub

'GENERIC
'A pool of messages
Public Function message(x As Integer, Optional specific1 As String, Optional specific2 As String) As Integer
Select Case x
    Case sure
        message = MsgBox("You are about to delete an entry" & specific1 _
        , vbOKCancel, "Confirmation")
    Case sure2
        message = MsgBox("You are about to delete a wood type and all the grades attached to it." _
        , vbExclamation + vbOKCancel, "Caution")
    Case clear_all
        message = MsgBox("You are about to delete all entries in this list. All current information will be lost. Proceed?" _
        , vbQuestion + vbYesNo, "Confirmation")
    Case Update
        message = MsgBox("You are about to change : " & specific1 & specific2 _
        , vbExclamation + vbOKCancel, "Caution")
    Case Save
        message = MsgBox("The data have changed, would you like to save before closing?" _
        , vbQuestion + vbYesNo, "Caution")

End Select
End Function

'GENERIC
'A pool of errors
Public Sub error(x As Integer, Optional Amendment As String)
Dim the_message As String
Select Case x
    Case Exists
        the_message = Amendment _
        + " already exists in the list."
        MsgBox the_message, vbExclamation, "Caution"
    Case notnumber
        the_message = "Some of the input was NOT numeric"
        MsgBox the_message, vbExclamation, "Caution"
    Case nonselect
        the_message = "Please, select an item from the list."
        MsgBox the_message, , "Unknown"
    Case overflow
        the_message = "Entries should be greater than zero and less than " & (Val(Amendment) + 0.001)
        MsgBox the_message, vbExclamation, "Caution"
    Case overchange
        the_message = "Changing the dimensions of the selection will effect its identity. Suggest to add the entry as new."
        MsgBox the_message, vbExclamation, "Suggestion"
    Case no_entry:
        the_message = "There is nothing to add"
        MsgBox the_message, vbExclamation, "Caution"
    Case empty_field:
        the_message = "Input should not be empty"
        MsgBox the_message, vbExclamation, "Caution"
    Case proper_size
         the_message = "The board cannot be, less than or, zero feet long. "
        MsgBox the_message, vbExclamation, "Caution"
    Case exceed_length
         the_message = "A force (reaction) cannot be applied, or span, beyond the length of the board."
        MsgBox the_message, vbExclamation, "Caution"
    Case noReactions
         the_message = "There should be at least " & Amendment & " Reaction(s) in the list."
        MsgBox the_message, vbExclamation, "Unable to comply"
    Case noEvaluate
         the_message = "There is nothing to evaluate."
        MsgBox the_message, vbExclamation, "Caution"
    Case listSupportFailure
         the_message = "There are existing Forces or Reactions (in the list)" & vbCr & "that apply beyond " & Amendment & " feet."
        MsgBox the_message, vbExclamation, "Unable to comply."
    Case noinclusion1
         the_message = "Calculations of axially loaded columns are not included in this version."
        MsgBox the_message, vbExclamation, "Unable to comply."
    Case noinclusion2
         the_message = "Calculations of combined axial and bending loads on columns are not included in this version."
        MsgBox the_message, vbExclamation, "Unable to comply."
    Case zeroMag
         the_message = "You are entering a load of (zero) magnitude."
        MsgBox the_message, vbExclamation, "Warning."
    Case noSpecies
         the_message = "You haven't selected Species or wood grade."
        MsgBox the_message, vbExclamation, "Warning."
    Case locationNonUnique
         the_message = "A reaction already exists, occupying the exast same location."
        MsgBox the_message, vbExclamation, "Warning."
    Case noLoads
         the_message = "There are no loads acting on the beam."
        MsgBox the_message, vbExclamation, "Unable to comply."
    Case MinusSpan
         the_message = "A negative span has no meaning."
        MsgBox the_message, vbExclamation, "Caution."
    Case Minusfield
         the_message = Amendment & " cannot be less than or equal to zero."
        MsgBox the_message, vbExclamation, "Unable to comply."

End Select
End Sub

' Child of frm_editsize
'Use the data in the matrix to populate the listview control
Public Sub extract_grades(to_list As ListView, from As matrix, species As String)
'Debug.Print "Attempting to extract"
Dim runner As Integer, run As Integer, one_line As String, population As Integer, words() As String
Dim itemx As ListItem
'clear the list first
to_list.ListItems.Clear
'Find the number of grades for the particular species
population = from.grade_count(species)
If (population > 1) Then
    For runner = 2 To population
        one_line = from.show_specific(species, runner)
        words = mdl_beam.fragment(one_line, Chr$(cutter))
        Set itemx = to_list.ListItems.Add()
            itemx.Text = words(1)
        For run = 2 To 7
            itemx.SubItems(run - 1) = words(run)
        Next
    Next
End If
End Sub

'Generic
'ensure that the data are not being dublicated
'This function performs either a single comparison check
'Or a dual (parallel) comparizon. It can compare two fields simultaneously
'This feature is useful to check for dublicate 'Reactions'
Public Function ItIsUnique(ToSearchFor As String, InThisListview As ListView, WhichColumn As Integer, Optional AndThis As String, Optional WhichColumn2 As Integer) As Boolean
Dim run As Integer
Dim TextFound As String
Dim TextFound2 As String

' Some error check
If (WhichColumn > InThisListview.ColumnHeaders.Count) Then
    ItIsUnique = False
    'call error access to subitems overflow
    Exit Function
End If

' Is the list empty?
If (InThisListview.ListItems.Count = 0) Then
    ItIsUnique = True
    Exit Function
End If

For run = 1 To InThisListview.ListItems.Count
    If (WhichColumn = 0) Then
        TextFound = InThisListview.ListItems(run).Text
    Else
        TextFound = InThisListview.ListItems(run).SubItems(WhichColumn)
    End If
    

    If (WhichColumn2 = 0 And AndThis <> "") Then
        TextFound2 = InThisListview.ListItems(run).Text
    ElseIf (WhichColumn2 > 0 And AndThis <> "") Then
        TextFound2 = InThisListview.ListItems(run).SubItems(WhichColumn)
    End If

    
    'remove any formating (for comparizon purposes) if the data is numeric
    If (IsNumeric(ToSearchFor) = True) Then
        ToSearchFor = Val(ToSearchFor)
        TextFound = Val(TextFound)
    End If
    If (IsNumeric(AndThis) = True) Then
        AndThis = Val(AndThis)
        TextFound2 = Val(TextFound2)
    End If

    

    If (TextFound = ToSearchFor And TextFound2 = AndThis) Then
        If (TextFound2 = "Reaction") Then
            ItIsUnique = False
            Exit Function
        ElseIf (TextFound2 = "Load") Then
            ItIsUnique = True
            Exit Function
        Else
            ItIsUnique = False
            Exit Function
        End If
    End If
Next
ItIsUnique = True

End Function

'Generic
'A function that automatically selects the text of a textBox
Public Sub selectText(InHere As TextBox)
InHere.SelStart = 0
InHere.SelLength = Len(InHere.Text)
InHere.SetFocus
End Sub

' This subroutine provides a way to handle the lists that appear in the forms of this project
' Instead of writting such code (to add delete or modify entries) in each form
Public Sub ListInterface(ThisList As ListView, theFields As Collection, command As Integer)

End Sub

