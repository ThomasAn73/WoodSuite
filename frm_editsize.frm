VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_editsize 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lumber-size Editor"
   ClientHeight    =   4770
   ClientLeft      =   3630
   ClientTop       =   3600
   ClientWidth     =   8205
   Icon            =   "frm_editsize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   8205
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   6240
      MaxLength       =   7
      TabIndex        =   6
      ToolTipText     =   "Weight, in pound per linear foot."
      Top             =   1080
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstsizes 
      Height          =   2055
      Left            =   360
      TabIndex        =   12
      Top             =   2040
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3625
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
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Index           =   4
      Left            =   6720
      TabIndex        =   11
      ToolTipText     =   "Close this editor"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Changes"
      Height          =   375
      Index           =   3
      Left            =   5400
      TabIndex        =   10
      ToolTipText     =   "Place the changes on file"
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add as new"
      Height          =   255
      Index           =   1
      Left            =   4920
      TabIndex        =   8
      ToolTipText     =   "Add the above information to the list."
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Remove"
      Height          =   255
      Index           =   2
      Left            =   6240
      TabIndex        =   9
      ToolTipText     =   "Delete the current selection from the list"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Modify"
      Height          =   255
      Index           =   0
      Left            =   3600
      TabIndex        =   7
      ToolTipText     =   "Amend the current selection"
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1200
      MaxLength       =   2
      TabIndex        =   1
      ToolTipText     =   "Nominal width."
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.000"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   2
      ToolTipText     =   "Actual thickness."
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   2760
      MaxLength       =   5
      TabIndex        =   3
      ToolTipText     =   "Actual width."
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   3600
      MaxLength       =   4
      TabIndex        =   4
      ToolTipText     =   "Adjustment factor (Cf)"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   4920
      MaxLength       =   4
      TabIndex        =   5
      ToolTipText     =   "Adjustment factor (Cfu)"
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   480
      MaxLength       =   2
      TabIndex        =   0
      ToolTipText     =   "Nominal thickness."
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "Flat Use Factor"
      Height          =   255
      Left            =   4920
      TabIndex        =   21
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Actual size"
      Height          =   255
      Left            =   2040
      TabIndex        =   20
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "d"
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label2 
      Caption         =   "b"
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   840
      Width           =   375
   End
   Begin VB.Line Line4 
      X1              =   360
      X2              =   360
      Y1              =   360
      Y2              =   1920
   End
   Begin VB.Line Line3 
      X1              =   7920
      X2              =   360
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line2 
      X1              =   7920
      X2              =   7920
      Y1              =   360
      Y2              =   1920
   End
   Begin VB.Line Line1 
      X1              =   360
      X2              =   7920
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Size factor (Cf)"
      Height          =   255
      Index           =   4
      Left            =   3600
      TabIndex        =   17
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Weight (lb/ft)"
      Height          =   255
      Index           =   3
      Left            =   6240
      TabIndex        =   16
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "d"
      Height          =   255
      Index           =   2
      Left            =   2880
      TabIndex        =   15
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "b"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   14
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Nominal size"
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
      Left            =   480
      TabIndex        =   13
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frm_editsize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const max_range = 999.999
Dim changed As Boolean

Private Sub Command1_Click(index As Integer)
Dim SelIndex As Integer
Dim IsSelect As Boolean

If (lstsizes.ListItems.Count > 0) Then
    SelIndex = lstsizes.SelectedItem.index
    IsSelect = lstsizes.SelectedItem.selected
End If
Select Case index
    Case 1      ' The ADD NEW command
        'Check the data to see if they are numerical
        If (check_valid_entries() = False) Then
            Exit Sub
        End If
        'Try to add the item if it doesn't already exist
        If (already_exists() = True) Then
            Call mdl_beam.error(err.Exists, "A " + Text1(0).Text + " X " _
                + Text1(1).Text & " or (" & Text1(1).Text & " X " & _
                Text1(0).Text & ") ")
            Exit Sub
        End If
        Call add_the_item
        'toggle the "save changes" flag
        changed = True
    Case 2      ' The DELETE CURRENT command
        'Is there anything selected??
        
        If (IsSelect = True) Then
            If (mdl_beam.message(ask.sure) = 1) Then
                lstsizes.ListItems.Remove (SelIndex)
                'toggle the "save changes" flag
                changed = True
            Else
                Exit Sub
            End If
        Else
            Call mdl_beam.error(err.nonselect)
        End If
     Case 0     ' The "Modify" command
        
        If (IsSelect = False) Then
            Call mdl_beam.error(err.nonselect)
            Exit Sub
        End If
        If (check_valid_entries() = False) Then
            Exit Sub
        End If
        
        If (same_dimension() = True) Then
            Call accept_change
            'toggle the "save changes" flag
            changed = True
        Else
            Call mdl_beam.error(err.overchange)
        End If
    Case 4      'Close button
        Call prompt_to_save
        Unload frm_editsize
    Case 3      'SAVE CHANGES button
        Call save_data
End Select


End Sub
Private Sub updateLumberSizes()
Dim between As Integer
Dim listRun As Integer
Dim oneLine As Integer

oneLine = lstsizes.ColumnHeaders.Count
ReDim Preserve LumberSizes(oneLine - 1, 0)
For listRun = 1 To lstsizes.ListItems.Count
    ReDim Preserve LumberSizes(oneLine - 1, listRun - 1)
    LumberSizes(0, listRun - 1) = Val(lstsizes.ListItems(listRun).Text)
    For between = 2 To oneLine
        LumberSizes(between - 1, listRun - 1) = Val(lstsizes.ListItems(listRun).SubItems(between - 1))
    Next
Next

End Sub

Private Function numeric() As Boolean
Dim temp As Integer
For temp = 0 To Text1.Count - 1
    If (IsNumeric(Text1(temp).Text) = False) Then
        numeric = False
        Exit Function
    End If
Next
numeric = True

End Function

Private Function add_the_item()
Dim itemx As ListItem
Dim run As Integer
Dim thisIndex As Integer

'make the addition to the lstsizes
Set itemx = lstsizes.ListItems.Add(properIndex(Text1(0), Text1(1)))
    itemx.Text = Text1(0).Text
For run = 1 To Text1.Count - 1
    itemx.SubItems(run) = Text1(run).Text
Next

End Function

Private Function properIndex(Test1 As String, test2 As String) As Integer
Dim run As Integer
Dim this1 As Integer
Dim this2 As Integer

For run = 1 To lstsizes.ListItems.Count
this1 = Val(lstsizes.ListItems(run).Text)
this2 = Val(lstsizes.ListItems(run).SubItems(1))
If (this1 >= Val(Test1) And this2 >= Val(test2)) Then
    properIndex = run
    Exit Function
End If
Next
properIndex = run
End Function


Private Sub form_Load()
Call init_lstsizes
Call load_list
End Sub

Private Sub init_lstsizes()

Dim columnX As ColumnHeader
Dim run As Integer
lstsizes.HideColumnHeaders = True

Set columnX = lstsizes.ColumnHeaders.Add
    columnX.width = Text1(0).width + (-lstsizes.Left + Text1(0).Left)
    'This one MUST be left aligned
    columnX.Alignment = lvwColumnLeft

For run = 1 To Text1.Count - 1
    Set columnX = lstsizes.ColumnHeaders.Add
        columnX.width = Text1(run).width + (Text1(run).Left - Text1(run - 1).Left - Text1(run - 1).width)
        If (run > 1) Then columnX.Alignment = lvwColumnRight
Next
End Sub

Private Sub updatefields()
Dim run As Integer
Dim SelIndex As Integer

SelIndex = lstsizes.SelectedItem.index

Text1(0).Text = Val(lstsizes.ListItems(SelIndex).Text)
For run = 1 To Text1.UBound
    Text1(run).Text = Val(lstsizes.ListItems(SelIndex).SubItems(run))
Next
End Sub

Private Sub format_fields()

Text1(0).Text = Format(Val(Text1(0).Text), "00")
Text1(1).Text = Format(Val(Text1(1).Text), "00")
Text1(2).Text = Format(Val(Text1(2).Text), "##0.00")
Text1(3).Text = Format(Val(Text1(3).Text), "##0.00")

Text1(4).Text = Format(Val(Text1(4).Text), "####0.000")
Text1(5).Text = Format(Val(Text1(5).Text), "####0.000")
Text1(6).Text = Format(Val(Text1(6).Text), "####0.000")

End Sub

Private Function range() As Boolean
Dim counter As Byte
Dim theValue As Single
For counter = 0 To (Text1.UBound)
    theValue = Val(Text1(counter).Text)
    If (theValue > max_range Or theValue <= 0) Then
        range = False
        Exit Function
    Else
        range = True
    End If
Next
End Function

Private Function check_valid_entries() As Boolean

If (numeric() = False) Then
    Call mdl_beam.error(err.notnumber)
    check_valid_entries = False
    Exit Function
End If
If (range() = False) Then
    Call mdl_beam.error(err.overflow, max_range)
    check_valid_entries = False
    Exit Function
Else
    Call format_fields
End If
    check_valid_entries = True
End Function

Private Sub accept_change()
Dim SelIndex As Integer

SelIndex = lstsizes.SelectedItem.index
lstsizes.ListItems(SelIndex).Text = Text1(0).Text
For run = 1 To Text1.Count - 1
    lstsizes.ListItems(SelIndex).SubItems(run) = Text1(run).Text
Next
lstsizes.SelectedItem.selected = False
End Sub

Private Function already_exists() As Boolean
Dim counter As Integer
Dim first As String
Dim second As String
Dim tempfirst As String
Dim tempsecond As String

' A 2X3 is the same as a 3X2
If (Text1(0).Text > Text1(1).Text) Then
    tempfirst = Text1(1).Text
    tempsecond = Text1(0).Text
Else
    tempfirst = Text1(0).Text
    tempsecond = Text1(1).Text
End If
For counter = 1 To lstsizes.ListItems.Count
    first = lstsizes.ListItems(counter).Text
    second = lstsizes.ListItems(counter).SubItems(1)
        
    If (tempfirst = first And tempsecond = second) Then
        already_exists = True
        Exit Function
    End If
Next
already_exists = False
End Function

'Check if the user has changed the name (dimention) of the item
Private Function same_dimension() As Boolean
Dim comp1 As String
Dim comp2 As String
Dim SelIndex As Integer

SelIndex = lstsizes.SelectedItem.index
comp1 = lstsizes.ListItems(SelIndex).Text
comp2 = lstsizes.ListItems(SelIndex).SubItems(1)
If (Text1(0).Text = comp1 And Text1(1).Text = comp2) Then
    same_dimension = True
    Exit Function
End If
same_dimension = False
End Function

Private Sub save_data()
Dim filetag As Integer, counter As Integer
Dim run As Integer
Dim oneLine As String

filetag = FreeFile      'Locate a free tag for the file
Open sizes_file For Output As #filetag
For counter = 1 To lstsizes.ListItems.Count
    oneLine = ""
    oneLine = oneLine & lstsizes.ListItems(counter) & Chr$(cutter)
    For run = 1 To Text1.UBound
        oneLine = oneLine & lstsizes.ListItems(counter).SubItems(run) & Chr$(cutter)
    Next
    oneLine = oneLine & vdcrlf
    Write #filetag, oneLine
Next
Close #filetag

'toggle the "save changes" flag
changed = False

Call updateLumberSizes
Call frmbeam.ListLumberSizes

End Sub

Private Sub load_list()
Dim reclaim As String
Dim counter As Integer, filetag As Integer
Dim theParts() As String
Dim itemx As ListItem
Dim run As Integer
'Dim thiskey As String

counter = 0
filetag = FreeFile
lstsizes.ListItems.Clear 'Clear the list (of possible impurities)
Open sizes_file For Input As #filetag
Do While (Not EOF(filetag))
    Input #filetag, reclaim
    theParts = mdl_beam.fragment(reclaim, " X" & Chr$(cutter), 7)
    
    If (Val(theParts(0)) <= Val(theParts(1))) Then
    'Add to the list
    Set itemx = lstsizes.ListItems.Add(properIndex(theParts(0), theParts(1)))
        itemx.Text = theParts(0)
    For run = 1 To Text1.Count - 1
        If (UBound(theParts) >= run) Then itemx.SubItems(run) = theParts(run)
    Next
    End If
Loop

Close #filetag
lstsizes.SortKey = 0
If (lstsizes.ListItems.Count > 0) Then lstsizes.SelectedItem.selected = False

End Sub

Private Sub lstsizes_Click()
Call updatefields
Call format_fields
End Sub

'Highlight the text when the textbox is in focous
Private Sub text1_GotFocus(index As Integer)
Select Case index
    Case 0 To Text1.Count - 1
        'set the cursor to the beginning of the textbox
        Text1(index).SelStart = 0
        'Highlight the text
        Text1(index).SelLength = Len(Text1(index).Text)
End Select
End Sub

Public Sub prompt_to_save()
If (changed = True) Then
    If (mdl_beam.message(Save) = yes) Then
        Call save_data
    End If
End If
Exit Sub
End Sub


