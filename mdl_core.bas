Attribute VB_Name = "mdl_core"
'This module contains functions that perform the necessary computation for beam design

'This array contains information on all vertical loads for each partition of the beam
Dim bendingLoads() As Double
'This array contains all axial loads
Dim axialLoads() As Double
'This array contains the magnitude of the bending moment at each partition
Dim MomentArray() As Double
'This array contains the magnitude of the shear at each partition
Dim shearArray() As Double

Public Function Evaluate(ThisList As ListView, beamLength As Integer, Optional Partitions As Integer = 100, Optional show As Boolean) As Boolean
Dim bendSuccess As Boolean
Dim dx As Double
Dim Reactions As Integer
'Magnitude of axial loads and bending loads
Dim AxialNum As Single
Dim BendingNum As Single

'differencial element (in feet)
dx = beamLength / Partitions

'the beam is divided into partitions of length "dx"
'information for each partition is stored in two arrays
ReDim bendingLoads(0 To Partitions)
ReDim axialLoads(0 To Partitions)

'is there anything in the list?
If (ThisList.ListItems.Count = 0) Then
    Call mdl_beam.error(err.noEvaluate)
    Evaluate = False
    Exit Function
End If

'calculate bending loads and axial loads for each partition of the beam
Call findLoads(ThisList, beamLength, Partitions, dx)

'find the number of axial and bending loads
AxialNum = ArraySum(axialLoads)
BendingNum = ArraySum(bendingLoads)

'Are there any loads?
If (AxialNum = 0 And BendingNum = 0) Then
    Call mdl_beam.error(err.noLoads)
    Evaluate = False
    Exit Function
End If

'Locate all reactions
Reactions = findreactions(ThisList, beamLength, Partitions, dx, AxialNum, BendingNum)

'the bending and axial arrays have some information
'based on which one has zero sum (or not) decide what proceedure to follow
If (AxialNum = 0) Then  'Proceed with design for BENDING

    'Are there enough reactions?
    'Need at least two for bending
    If (Reactions < 2) Then
        Call mdl_beam.error(err.noReactions, "two")
        Evaluate = False
        Exit Function
    End If
    'Calculate the moments and shear
    'Also return the values for max moment and Max shear
    currspecs.MaxMoment = bendingMoments(Partitions, dx, show)
    currspecs.MaxShear = shear(Partitions, dx, show)
    
    'Find the required number of beams
    Call frmbeam.ShowNumBeams
    
    'Present the results in a graphical form
    Call ImageControl(Partitions, show)
    
    Evaluate = True

ElseIf (BendingNum = 0) Then
    'There is axial loading without bending
    'Proceed with design for (simple) axialy loaded column
    
    'Need at least one reaction in the axial direction
    If (Reactions < 1) Then
        Call mdl_beam.error(err.noReactions, "one axial")
        Evaluate = False
        Exit Function
    End If
    If (currspecs.compression = True) Then
        Call comprNoBend(AxialNum, 1)
    Else
        Call TensionNoBend
    End If
Else
    'There is axial load and bending
    If (currspecs.compression = True) Then
        Call comprYesBend
    Else
        Call TensionYesBend
    End If

End If

End Function

'This function handles calculations of tension without bending forces
Private Sub TensionNoBend()
'The wet service factor for tension (Cm_Ft) is always=1.0
Dim run As Integer
Dim FtStar As Double
Dim reqArea As Double
Dim numberNeeded As Integer
Dim RunTwice As Integer

For run = 0 To UBound(LumberSizes, 2)
    'actual thickness
    currspecs.beamThickness = LumberSizes(2, run)
    'actual width
    currspecs.beamWidth = LumberSizes(3, run)
    'The area is
    currspecs.Area = currspecs.beamThickness * currspecs.beamWidth
    numberNeeded = 0
    For RunTwice = 1 To 2
        If (numberNeeded <= 1) Then
            FtStar = currspecs.FsubT * correctionFactors(run, False, False, , False, True, False)
        Else
            FtStar = currspecs.FsubT * correctionFactors(run, , False, , False, True, False)
        End If
        reqArea = currspecs.UnbracedLengthLoad / FtStar
    
        numberNeeded = roundUp(reqArea / currspecs.Area)
        'No need to repeat if the result is already: 'one'
        If (numberNeeded <= 1) Then Exit For
    Next
    'Output that number on the lumberSizes list
    frmbeam.lstsizes.ListItems(run + 1).ListSubItems(3).Text = numberNeeded

Next
End Sub

Private Sub TensionYesBend()
Call mdl_beam.error(err.noinclusion2)

End Sub

'This function is coordinating the graphing of shear and bending diagrams
Private Sub ImageControl(Partitions As Integer, show As Boolean)

If (show = False) Then
    Unload frm_graph
    Exit Sub
End If

'The Shear
frm_graph.TextShear.Text = Format(currspecs.MaxShear, "#######.000")
Call SketchIt(frm_graph.TheShear, shearArray, Partitions, currspecs.MaxShear)
    
'The Moment
frm_graph.TextMoment.Text = Format(currspecs.MaxMoment, "#######.000")
Call SketchIt(frm_graph.picBox, MomentArray, Partitions, currspecs.MaxMoment)
frm_graph.show

End Sub

'Compression (or tension) without bending
Private Sub comprNoBend(appliedLoad As Single, endCond As Integer)
'if we are here there are no bending loads acting on the beam
'the "axialLoads" array is already updated
'we will calculate the allowable compression for every size of lumber
Dim run As Integer
Dim sl As Double 'slenderness ratio
Dim Fce As Double
Dim Kce As Double
Dim FcStar As Double
Dim FcPrime As Double
Dim Cp As Double
Dim theC As Double
Dim Pmax As Double
Dim numberNeeded As Integer
Dim RunTwice As Integer

For run = 0 To UBound(LumberSizes, 2)
    'actual thickness
    currspecs.beamThickness = LumberSizes(2, run)
    'actual width
    currspecs.beamWidth = LumberSizes(3, run)
    'The area is
    currspecs.Area = currspecs.beamThickness * currspecs.beamWidth
    'Compute the slenderness ratio
    'The sixth item of the list is the Ke factor
    
    sl = frmbeam.lstFactors.ListItems(6).SubItems(2) * currspecs.MaxUnbracedLength * 12 / currspecs.beamThickness
    Kce = frmbeam.lstFactors.ListItems(8).SubItems(2)
    Fce = Kce * currspecs.ElastMod * correctionFactors(run, False, False, False, True, False, True) / (sl ^ 2)
    theC = frmbeam.lstFactors.ListItems(7).SubItems(2)
    numberNeeded = 0
    
    For RunTwice = 1 To 2
    'provision for the Cr factor
    If (numberNeeded <= 1) Then
        'do not include Cr factor
        FcStar = currspecs.FsubC * correctionFactors(run, False, False, , , , True)
    Else
        'Include the repetitive member factor
        FcStar = currspecs.FsubC * correctionFactors(run, True, False, , , , True)
    End If
    Cp = (1 + Fce / FcStar) / (2 * theC) - Sqr(((1 + Fce / FcStar) / (2 * theC)) ^ 2 - (Fce / FcStar) / theC)

    FcPrime = FcStar * Cp
    Pmax = FcPrime * currspecs.Area
    
    'divide the applied load with Pmax to find the number of beams needed
    numberNeeded = roundUp(currspecs.UnbracedLengthLoad / Pmax)
    If (numberNeeded <= 1) Then Exit For
    
    'We need to loop once more, to account for the repetitive member factor
    'Loop While (numberNeeded = 0)
    Next
    
    'Output that number on the lumberSizes list
    frmbeam.lstsizes.ListItems(run + 1).ListSubItems(3).Text = numberNeeded
    
Next


'Call mdl_beam.error(err.noinclusion1)
End Sub

'Compression combined with bending
Private Sub comprYesBend()
'..
Call mdl_beam.error(err.noinclusion2)
End Sub

'Utility function that sums the elements of any arbitrary array
Public Function ArraySum(thisArray() As Double) As Double
Dim run As Integer

For run = LBound(thisArray) To UBound(thisArray)
    ArraySum = ArraySum + thisArray(run)
Next
End Function

'Traverse the list, locate all loads and fill the "master" array
Private Sub findLoads(wholeList As ListView, beamLength As Integer, Partitions As Integer, dx As Double)

Dim orientation As Integer  'a multiplier
Dim atPoint As Integer   'holds the current insertion point of the array
Dim theSpan As Integer
Dim run1 As Integer, listRun As Integer
Dim listEntry As ListItems, listSubEntry As ListSubItems

'traverse the list
Set listEntry = wholeList.ListItems

For listRun = 1 To listEntry.Count
    Set listSubEntry = listEntry(listRun).ListSubItems
    orientation = direction(wholeList, listRun)

    If (listEntry(listRun).Text = "Load") Then
        'point load
        If (Val(listSubEntry(L.span)) = 0) Then
            'filter between axial and bending loads
            If (listSubEntry(L.direction) = "Up" Or listSubEntry(L.direction) = "Down") Then
                atPoint = listSubEntry(L.location) / dx
                bendingLoads(atPoint) = bendingLoads(atPoint) + orientation * listSubEntry(L.magnitude)
            Else 'place the axial loads in a separate array
                atPoint = listSubEntry(L.location) / dx
                axialLoads(atPoint) = axialLoads(atPoint) + orientation * listSubEntry(L.magnitude)

            End If
        Else  'Uniform load
            theSpan = listSubEntry(L.span) / dx
            
            'distribute to several partitions
            For run1 = 0 To theSpan - 1
                atPoint = listSubEntry(L.location) / dx + run1
                bendingLoads(atPoint) = bendingLoads(atPoint) + orientation * listSubEntry(L.magnitude) * dx
            Next
        End If
    End If
Next

End Sub

'locate all reactions, find their magnitude and place them in the "R_locations" array
'reaction should appear as negative
'this function returns the number of reactions found
Private Function findreactions(wholeList As ListView, beamLength As Integer, Partitions As Integer, dx As Double, axialL As Single, bendingL As Single) As Integer

Dim R_Location() As Integer
Dim here As Integer   'holds the current insertion point of an array
Dim listRun As Integer, run As Integer       'counters
Dim listEntry As ListItems, listSubEntry As ListSubItems

currspecs.MaxUnbracedLength = 0
Set listEntry = wholeList.ListItems
here = 0

For listRun = 1 To listEntry.Count
    Set listSubEntry = listEntry(listRun).ListSubItems
    'Isolate all reactions and put their locations inside an array
    If (listEntry(listRun).Text = "Reaction") Then
        
        ReDim Preserve R_Location(here)
        R_Location(here) = listSubEntry(L.location) / dx
        here = here + 1
    End If
Next

'Sort reactions according to ascenting order of location
If (here > 1) Then Call sortThis(R_Location())
'Exit if there are no reactions
If (here < 1) Then
    findreactions = 0
    Exit Function
End If

If (here = 2 And axialL = 0) Then
    'bending with two reactions
    Call Two_Reactions(R_Location, Partitions, dx)
ElseIf (here > 2 And axialL = 0) Then
    'bending with multiple spans
    Call threeMomentEq(R_Location, Partitions, dx)
ElseIf (bendingL = 0 And axialL <> 0) Then
    'There is compression (or tension) without bending
    Call ReactionNoBend(R_Location, Partitions, dx, axialL)
ElseIf (bendingL <> 0 And axialL <> 0) Then
    'there is compression (or tension) and bending
    
End If

findreactions = here
End Function

'Finds the reactions for a column without bending
Private Sub ReactionNoBend(TheLocations() As Integer, thePartitions As Integer, dx2 As Double, columnLoad As Single)
'Assamptions:
'1)A single load will transfer to the last reaction
'2)Intermediate reactions have a value of zero, they only act as bracings
'3)for the span where the load is acting we will assume the Ke value specified by the user
'   however, for all other spans (with no direct application of load) we will assume pinned supports Ke=1.0

Dim run As Integer

'Transfer the load to the last reaction (towards the compression side)
If (columnLoad > 0) Then
    'axialLoads(TheLocations(UBound(TheLocations))) = -columnLoad
    Call frmbeam.ListUpdate("Reaction", TheLocations(UBound(TheLocations)) * dx2, -columnLoad, True)
    Call unbracedLength(TheLocations, thePartitions, dx2, True)

ElseIf (columnLoad < 0) Then
    'axialLoads(TheLocations(LBound(TheLocations))) = -columnLoad
    Call frmbeam.ListUpdate("Reaction", TheLocations(LBound(TheLocations)) * dx2, -columnLoad, True)
    Call unbracedLength(TheLocations, thePartitions, dx2, False)

End If


End Sub

'Find the maximum unbraced length
'and the magnitude of the load applied to that length
'also, find whether that section is in tension or compression
Private Sub unbracedLength(Locations() As Integer, thePartitions As Integer, dx2 As Double, rightwards As Boolean)
Dim run As Integer
Dim runIn As Integer
Dim accumulate As Single
Dim between As Single
Dim upper As Integer
Dim LoadAt As Integer
Dim OuterReaction As Integer

upper = UBound(Locations)
If (rightwards = True) Then
    OuterReaction = Locations(UBound(Locations))
Else
    OuterReaction = Locations(LBound(Locations))
End If
'treat the loads as if the beam is braced at those locations
For run = 0 To UBound(axialLoads)
    If (axialLoads(run) <> 0) Then
        upper = upper + 1
        ReDim Preserve Locations(upper)
        Locations(upper) = run
    End If
Next
Call sortThis(Locations)

between = 0
accumulate = 0
'rightwards: tells us which way is the ground support
'the 'offset' is used to flip the direction of the loops
If (rightwards = True) Then
    offset = 0
Else
    offset = UBound(Locations)
End If
'measure the distance of each span and find the longest
'run through each span and find the magnitude the loads applied to it

'start at second (or second to last) location
For run = (1 - offset) To (UBound(Locations) - offset)
    'distance between second and first location
    'distance between second to last and last location (for reverse loop)
    between = Abs(Locations(Abs(run)) - Locations(Abs(run - 1))) * dx2
    If (between > currspecs.MaxUnbracedLength) Then
        currspecs.MaxUnbracedLength = between
        For runIn = Locations(Abs(run - 1)) To (Locations(Abs(run)) - 1 * Sgn(Abs(run) - Abs(run - 1))) Step Sgn(Abs(run) - Abs(run - 1))
            accumulate = accumulate + axialLoads(runIn)
            currspecs.UnbracedLengthLoad = accumulate
            If (axialLoads(runIn) <> 0) Then LoadAt = runIn
        Next
    End If
Next
'Account for the last partition
accumulate = accumulate + axialLoads(runIn)
currspecs.UnbracedLengthLoad = accumulate
If (axialLoads(runIn) <> 0) Then LoadAt = runIn


'determine if the section is under compression or tension
If ((currspecs.UnbracedLengthLoad > 0 And LoadAt < OuterReaction) Or _
    (currspecs.UnbracedLengthLoad <= 0 And LoadAt > OuterReaction)) Then
    'we have compression
    currspecs.compression = True
Else
    'We have tension
    currspecs.compression = False
End If
'Make sure the value of the load is positive, for the sake of calculation (without glitches)
currspecs.UnbracedLengthLoad = Abs(currspecs.UnbracedLengthLoad)

End Sub

Private Sub threeMomentEq(Locations() As Integer, thePartitions As Integer, dx2 As Double)
'bendingLoads() is a global array

Dim run As Integer, run2 As Integer 'Utility counters

Dim Matrix_3Momts() As Double 'contains data of the three moment equation
Dim TotalColumns As Integer
Dim TotalRows As Integer

'Variables used for the elements of the 3 Moment eq.
Dim L1 As Double, L2 As Double
Dim rightSide As Double
Dim sum1 As Double, sum2 As Double
Dim a1 As Double, b1 As Double

Dim overhung1 As Double, overhung2 As Double
Dim multiplier As Double
Dim momentsAt As Double
Dim oneReaction As Double
Dim toHere As Integer

'number of rows for Matrix_3Momts
'It is 2 less, than the number of reactions
'locations array is from (0 to upperbound)
TotalRows = UBound(Locations) - 1

'number of columns for Matrix_3Momts
'the three moment equation will generate as many data as the number of reactions
TotalColumns = UBound(Locations) + 1

'instantiate Matrix_3Momts
'Give an extra column for the right side of the 3Moment eq.
'Give 2 extra rows of zeros that represent the zero moment at the first and last reaction
ReDim Matrix_3Momts(1 To TotalColumns + 1, 0 To TotalRows + 1)

'the Matrix_3momts will be processed using the Gauss-Jordan reduction method
'Using rules of matrix algebra, zeros will be produced
' in the upperright and lower left corners of the matrix
' then division with proper constants will yield a aubmented unit matrix
' the last column of that matrix would contain the moment at each reaction

'totalRows are the working rows, because remember that there are 2 extra rows added
'   one at the beginning and one at the end of the matrix
For run = 1 To TotalRows
    'enter each element (left side of 3Moment eq) into the matrix
    ' Ma*L1+ Mb*2*(L1+L2) + Mc*L2
    
    'locations() array contains the locations of the reactions in number of partitions
    'NOT in feet!! So we need to multiply with dx2
    L2 = (Locations(run + 1) - Locations(run)) * dx2
    L1 = (Locations(run) - Locations(run - 1)) * dx2
    Matrix_3Momts(run, run) = L1
    Matrix_3Momts(run + 1, run) = 2 * (L1 + L2)
    Matrix_3Momts(run + 2, run) = L2
    
    'Calculate the right side of the three moment equation
    'using the point load formula
    '= -SUM[ (P1*a1/L1)*(L1^2-a1^2) ] - SUM[ (P2*b2/L2)*(L2^2-b2^2)]
    
    'This is the first sum (Span A)
    For run2 = Locations(run - 1) To Locations(run)
        a1 = (run2 - Locations(run - 1)) * dx2
        sum1 = sum1 - bendingLoads(run2) * a1 / L1 * (L1 ^ 2 - a1 ^ 2)
    Next    'Iteration that performs summation
    
    'This is the second sum (Span B)
    For run2 = Locations(run) To Locations(run + 1)
        b2 = (Locations(run + 1) - run2) * dx2
        sum2 = sum2 - bendingLoads(run2) * b2 / L2 * (L2 ^ 2 - b2 ^ 2)
    Next    'Iteration that performs summation
    
    'How to deal if the beam has overhung
    'The first and last moments are no longer Zero
    If (run = 1) Then
        overhung1 = about(Locations(0), dx2, 0) * Locations(0) * dx2
        sum1 = sum1 - overhung1
        If (Locations(0) <> 0) Then Matrix_3Momts(TotalColumns + 1, 0) = overhung1 / (Locations(0) * dx2)
    End If
    If (run = TotalRows) Then
        overhung2 = -about(Locations(run + 1), dx2, thePartitions) * (Locations(run + 1) - Locations(run)) * dx2
        sum2 = sum2 - overhung2
        If ((Locations(run + 1) - Locations(run)) <> 0) Then Matrix_3Momts(TotalColumns + 1, run + 1) = overhung2 / ((Locations(run + 1) - Locations(run)) * dx2)
    End If
    
    'place the result into the last column of Matrix_3Momts
    Matrix_3Momts(TotalColumns + 1, run) = sum1 + sum2
    
    'initialise the sums for the next iteration
    sum1 = 0
    sum2 = 0
Next    'iteration that generates equations
'The preliminary DATA have now been entered into matrix_3momts

'THIS IS THE MATRIX ALGEBRA PORTION
'do it if there are enough rows to work with
If (TotalRows > 1) Then
    
    'produce zeros in the lower left corner of the matrix
    'Start from the second row
    For run = 2 To TotalRows
        
        ' a non zero element
        If (Matrix_3Momts(run, run) <> 0) Then
            multiplier = Matrix_3Momts(run, run - 1) / Matrix_3Momts(run, run)
            'subtract current row from the previous row
            For run2 = 1 To TotalColumns + 1
                'and place the result in the current row
                Matrix_3Momts(run2, run) = Matrix_3Momts(run2, run - 1) - multiplier * Matrix_3Momts(run2, run)
            Next
        End If
    Next
 
    'produce zeros in the upper right corner of the matrix
    'start at the second to last row
    For run = TotalRows - 1 To 1 Step -1
        'a non zero element
        If (Matrix_3Momts(run + 2, run) <> 0) Then
            ' run+2 is the proper column
            multiplier = Matrix_3Momts(run + 2, run + 1) / Matrix_3Momts(run + 2, run)
            'subtract current row from the one bellow
            For run2 = 1 To TotalColumns + 1
                'and place the result in the current row
                Matrix_3Momts(run2, run) = Matrix_3Momts(run2, run + 1) - multiplier * Matrix_3Momts(run2, run)
            Next
        End If
    Next
   
    'divide with proper constant to create unit matrix
    For run = 1 To TotalRows
        Matrix_3Momts(TotalColumns + 1, run) = Matrix_3Momts(TotalColumns + 1, run) / Matrix_3Momts(run + 1, run)
        Matrix_3Momts(run + 1, run) = 1
        'Debug.Print "MOMENTS ??=="; Int(Matrix_3Momts(TotalColumns + 1, run))
    Next
'The moments have been found

'in case of simple two-span beam, do the following
ElseIf (TotalRows = 1) Then 'solve the trivial form
    'we only have one row and 4 columns
    multiplier = Matrix_3Momts(2, 1)
    For run = 1 To TotalColumns + 1
        Matrix_3Momts(run, 1) = Matrix_3Momts(run, 1) / multiplier
    Next
End If
'==========

'show me
'    Debug.Print "---------------------"
'    For run = 0 To TotalRows + 1
'        For run2 = 1 To TotalColumns + 1
'            Debug.Print Int(Matrix_3Momts(run2, run));
 '       Next
 '       Debug.Print vbCrLf
 '   Next
    
'calculate the  reactions and place them in the master array
For run = 1 To UBound(Locations)
    'Moment from current position all the way to the start
    momentsAt = about(Locations(run), dx2, 0)
    'Moments about a point 2 to find the reaction at some point 1
    
    '(keep this)
    oneReaction = (momentsAt - Matrix_3Momts(TotalColumns + 1, run)) / ((Locations(run) - Locations(run - 1)) * dx2)
    
    bendingLoads(Locations(run - 1)) = bendingLoads(Locations(run - 1)) + oneReaction
    'Debug.Print "ONE REACTION is......"; oneReaction; vbCrLf
    Call frmbeam.ListUpdate("Reaction", Locations(run - 1) * dx2, oneReaction, False)

Next

'and the last reaction
'Second to last reaction all the way to the end
run = run - 2
momentsAt = about(Locations(run), dx2, thePartitions)

'momentsAt = momentsAt + about(locations(run), dx2, 0)
oneReaction = -(momentsAt + Matrix_3Momts(TotalColumns + 1, run)) / ((Locations(run + 1) - Locations(run)) * dx2)
bendingLoads(Locations(run + 1)) = bendingLoads(Locations(run + 1)) + oneReaction
Call frmbeam.ListUpdate("Reaction", Locations(run + 1) * dx2, oneReaction, False)
'Debug.Print "ONE REACTION is......"; oneReaction; vbCrLf


End Sub

'find moments about a point "A"
Private Function about(pointA As Integer, dx2 As Double, Optional ToPointB As Integer) As Double
Dim run As Integer
Dim thisWay As Integer
If ((pointA - ToPointB) > 0) Then
    thisWay = -1
Else
    thisWay = 1
End If
For run = pointA To ToPointB Step thisWay
    about = about + thisWay * bendingLoads(run) * Abs(pointA - run) * dx2
Next
End Function

'Finding reactions for a bending load situation
Private Sub Two_Reactions(Locations() As Integer, thePartitions As Integer, diffElement As Double)
'bendingLoads()  is a global
Dim here As Integer
Dim R_distance As Double
Dim reaction As Double
Dim momentsAbout1 As Double, sumOfLoads As Double
Dim orientation As Integer  'a multiplier

'calculate reactions (assuming there are only two)
For here = 0 To thePartitions
    'sum of moments about R1
    momentsAbout1 = momentsAbout1 + diffElement * (bendingLoads(here) * (here - Abs(Locations(0))))
    'sum of all loads
    sumOfLoads = sumOfLoads + bendingLoads(here)
Next
'Debug.Print "Moments about R1..."; momentsAbout1

R_distance = diffElement * Abs(Locations(1) - Locations(0))
reaction = -1 * momentsAbout1 / R_distance
'add the value of the reaction in the "master" array
bendingLoads(Locations(1)) = bendingLoads(Locations(1)) + reaction
'Place the magnitude of the reaction into the list
Call frmbeam.ListUpdate("Reaction", Locations(1) * diffElement, reaction, False)

'use the same variable to calculate next reaction
reaction = -1 * (sumOfLoads + reaction)
'add the value of the reaction in the "master" array
bendingLoads(Locations(0)) = bendingLoads(Locations(0)) + reaction
Call frmbeam.ListUpdate("Reaction", Locations(0) * diffElement, reaction, False)

End Sub

'Bubble sort
'GENERIC
Private Sub sortThis(SortMe() As Integer)
Dim temp As Integer
Dim run As Integer
Dim runOut As Integer

For runOut = 1 To UBound(SortMe)
    For run = 1 To UBound(SortMe) - runOut + 1
        temp = SortMe(run - 1)
        If (SortMe(run - 1) > SortMe(run)) Then
            SortMe(run - 1) = SortMe(run)
            SortMe(run) = temp
        End If
    Next
Next
End Sub

'Translate direction into a positive or negative multiplier
'It can be placed in a function
Private Function direction(inlist As ListView, item As Integer)
Select Case inlist.ListItems(item).SubItems(L.direction)
    Case "Up", "Left"
        direction = -1
    Case "Down", "Right"
        direction = 1
End Select

End Function

'Perform the standard moment calculation
Private Function bendingMoments(Partitions As Integer, dx As Double, showit As Boolean) As Double
Dim MaxMoment As Single
Dim maxAt As Single
Dim TheDistance As Double
Dim TheMoment As Double
Dim sumTheMoments As Double
Dim thisPoint As Integer, between As Integer, ex As Integer
Dim yes As Byte
Dim drawWhere As PictureBox
Dim theColor As Long

ReDim MomentArray(Partitions)
MaxMoment = 0

Set drawWhere = frm_graph.picBox

'okey, now do the moment calculations
For thisPoint = 0 To Partitions
    sumTheMoments = 0
    'from current location to the beginning
    For between = thisPoint To 0 Step -1
        TheDistance = (between - thisPoint) * dx
        TheMoment = bendingLoads(between) * TheDistance
        sumTheMoments = sumTheMoments + TheMoment
    Next
    If (Abs(sumTheMoments) > Abs(MaxMoment)) Then
        MaxMoment = sumTheMoments
        maxAt = thisPoint * dx
    End If
    MomentArray(thisPoint) = sumTheMoments
Next

currspecs.MaxMomentLoc = maxAt
'If (showit = True) Then
'    frm_graph.picBox.Cls
'    theColor = RGB(70, 79, 173)
'    For thisPoint = 0 To Partitions
'        'make a graphical representation
'        Call frm_graph.draw(drawWhere, thisPoint, MomentArray(thisPoint), Partitions * 1, Abs(MaxMoment) * 2.5, theColor)
'    Next

'End If

'frm_graph.TextMoment.Text = Format(MaxMoment, "#######.000")

'return the value
bendingMoments = MaxMoment
End Function

Private Function shear(thePartitions As Integer, dx2 As Double, showit As Boolean) As Double

Dim run As Integer
Dim between As Integer
Dim sumForces As Double
Dim MaxShear As Double
Dim maxAt As Double
Dim drawWhere As PictureBox
Dim theColor As Long

ReDim shearArray(thePartitions)

Set drawWhere = frm_graph.TheShear

For run = 0 To thePartitions
sumForces = 0
    For between = run To 0 Step -1
        sumForces = sumForces - bendingLoads(between)
    Next
shearArray(run) = sumForces

If (Abs(sumForces) > Abs(MaxShear)) Then
    maxAt = run * dx2
    MaxShear = sumForces
End If
Next


currspecs.MaxShearLoc = maxAt
'frm_graph.TextShear.Text = Format(MaxShear, "#######.000")

'If (showit = True) Then
'frm_graph.TheShear.Cls
'theColor = RGB(70, 79, 173)
'    For run = 0 To thePartitions
'        Call frm_graph.draw(drawWhere, run, shearArray(run), thePartitions * 1, Abs(MaxShear) * 2.5, theColor)
'    Next
'End If

'return the value
shear = MaxShear

End Function

'This function recieves the species name and the grade and searces for the Fb, Fc and Ft values
Public Sub determine_FsubB(ThisSpecies As String, ThisGrade As String)
Dim TheData() As String
Dim oneLine As String
Dim run As Integer
'See if the information is valid
If (ThisSpecies = "" Or ThisGrade = "") Then Exit Sub

'we need to search the tree matrix to find the proper value
'The first line is the name of the species, so start from the second line
For run = 2 To tree.grade_count(ThisSpecies)
    oneLine = tree.show_specific(ThisSpecies, run)
    TheData = mdl_beam.fragment(oneLine, Chr$(cutter))
    'Element "0" is the species name (again)
    'Element "1" is the grade name
    'Element "2" is Fb
    'Element "3" is Ft
    'Element "7" is the modulus of elasticity
    If (TheData(1) = ThisGrade) Then
        currspecs.FsubB = Val(TheData(2))
        currspecs.FsubC = Val(TheData(6))
        currspecs.FsubT = Val(TheData(3))
        currspecs.ElastMod = Val(TheData(7))
        Exit Sub
    End If
Next

If (currspecs.FsubB = 0) Then
'some error
End If

End Sub

'Child of frmbeam
'It produces the result number of boards needed (to be printed)
Public Function how_many(C_factors As Double) As Integer
Dim TheNumber As Single

Dim momentFromWeight As Double
Dim denominator As Double

'currspecs.beamLength = Val(frmbeam.txt_Blength.Text)
TheNumber = 0
'C_factors = correctionFactors(True, True, True, True, True)
'Debug.Print currspecs.MaxUnbracedLength
'Debug.Print currspecs.MaxMomentLoc

'Mw= w*L^2/8
momentFromWeight = currspecs.beamWeight * currspecs.beamLength ^ 2 / 8

'Calculate the requirement
'N=Mm/(S*Fb*Cf/12-Mw)
denominator = (currspecs.SectMod / 12 * (currspecs.FsubB * C_factors) - momentFromWeight)
If (currspecs.FsubB <> 0) Then TheNumber = Abs(currspecs.MaxMoment) / denominator

how_many = roundUp(TheNumber)
'Negative result means that the particular beam is an overkill
'(moment from the weight of the beam is larger than the max moment)
If (how_many < 1) Then how_many = 1

End Function

'generic
'rounds a number up to the nearest integer
'There is an option for negative numbers to not be accepted
Private Function roundUp(ThisNumber As Variant, Optional acceptNegative As Boolean = False) As Integer
Dim test As Double

If (Abs(Val(ThisNumber)) > 999) Then
    roundUp = -1
    Exit Function
End If
test = Val(ThisNumber) - Int(Val(ThisNumber))

If (test >= 0) Then
    roundUp = Int(Val(ThisNumber)) + 1
ElseIf (test < 0) Then
    roundUp = Int(Val(ThisNumber))
End If
If (acceptNegative = False And test < 0) Then
    roundUp = 1
    Exit Function
End If

End Function

'correction factors as needed for bending or compression
'(True) means that the particular factor will be included in computation
'User choices also, influence whether a factor will be included in computations
Public Function correctionFactors(lumberIndex As Integer, Optional Cr As Boolean = True, Optional Cfu As Boolean = True, Optional Cf As Boolean = True, Optional Cm As Boolean = True, Optional Cd As Boolean = True, Optional compression As Boolean = False) As Double
'lumberindex is used for the flat use and the size factors
'it is the index (within "lumbersizes" array) of the lumber size that is being evaluated at this moment
correctionFactors = 1

If (Cr = True) Then
    'Repetitive member factor is included for lumber of less than 4 inches
    If ((currspecs.beamWidth <= 4 Or currspecs.beamThickness <= 4) And c.r > 0) Then
        c.r = Val(lstFactors.ListItems(5).SubItems(2))
        correctionFactors = correctionFactors * c.r
    End If
End If

If (Cfu = True) Then
    'Flat use factor (Variable)
    If (factor.flat = True) Then
        c.fu = LumberSizes(5, lumberIndex)
        correctionFactors = correctionFactors * c.fu
    End If
End If

If (Cf = True) Then
    If (compression = False) Then
    'size factor (variable)
    c.f = LumberSizes(4, lumberIndex)
    correctionFactors = correctionFactors * c.f
    Else
        Select Case currspecs.beamWidth
            Case 0 To 4
                c.f_Fc = 1.15
            Case 5 To 6
                c.f_Fc = 1.1
            Case 7 To 8
                c.f_Fc = 1.05
            Case 9 To 12
                c.f_Fc = 1
            Case Else
                c.f_Fc = 0.9
        End Select
    correctionFactors = correctionFactors * c.f_Fc
    End If
End If

If (Cm = True) Then
    
    If (factor.wet = True) Then
        'Wet service factor for bending
        If (currspecs.FsubB * c.f <= 1150) Then
            c.m = (frmbeam.lstFactors.ListItems(2).SubItems(2))
        Else
            c.m = 1
        End If
        'Wet service factor for compression
        If (currspecs.FsubC * c.f_Fc <= 750) Then
            c.m_Fc = 0.8
        Else
            c.m_Fc = 1
        End If
        
        If (compression = False) Then
            correctionFactors = correctionFactors * c.m
        Else
            correctionFactors = correctionFactors * c.m_Fc
        End If
    End If
End If

If (Cd = True) Then
    'Load duration factor
    If (factor.duration = True) Then
        c.d = frmbeam.lstFactors.ListItems(3).SubItems(2)
        correctionFactors = correctionFactors * c.d
    End If
End If

End Function

Public Sub ZeroThefactors()

With currspecs
    .Area = 0
    .beamLength = 0
    .beamThickness = 0
    .beamWeight = 0
    .beamWidth = 0
    .ElastMod = 0
    .FsubB = 0
    .MaxMoment = 0
    .MaxShear = 0
    .MaxUnbracedLength = 0
    .SectMod = 0
End With
End Sub

Private Sub SketchIt(drawWhere As PictureBox, thisArray() As Double, Partitions As Integer, MaxY As Double)
Dim theColor As Long
Dim run As Integer

drawWhere.Cls
theColor = RGB(70, 79, 173)
    For run = 0 To Partitions
        Call frm_graph.draw(drawWhere, run, thisArray(run), Partitions * 1, Abs(MaxY) * 2.5, theColor)
    Next

End Sub
