VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ld_stack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim load As Collection
Dim reaction As Collection
Dim brace As Collection

Public Sub load_push(loadname As String)
load.Add loadname
'Debug.Print "Pushing "; loadname
End Sub
Public Function load_count() As Integer
load_count = load.Count
End Function
Public Function load_pop() As String
load_pop = load.Item(load.Count)
'Debug.Print "Popping "; load.Item(load.Count)

load.Remove (load.Count)
End Function

Public Sub reaction_push(reactionName As String)
reaction.Add reactionName
End Sub
Public Function reaction_count() As Integer
reaction_count = reaction.Count
End Function
Public Function reaction_pop() As String
reaction_pop = reaction.Item(reaction.Count)
reaction.Remove (reaction.Count)
End Function


Public Sub brace_push(braceName As String)
brace.Add braceName
End Sub

Public Function brace_count() As Integer
brace_count = brace.Count
End Function

Public Function brace_pop() As String
brace_pop = brace.Item(brace.Count)
brace.Remove (brace.Count)
End Function

Private Sub Class_Initialize()
Set load = New Collection
Set reaction = New Collection
Set brace = New Collection
End Sub

Public Sub pop_all()
Call clear_loads
Call clear_reactions
Call clear_braces
End Sub

Private Sub clear_loads()
Dim run As Integer
Dim population As Integer
population = load.Count
If (population > 0) Then
    For run = population To 1 Step -1
        load.Remove (run)
    Next
End If
End Sub

Private Sub clear_reactions()
Dim run As Integer
Dim population As Integer
population = reaction.Count
If (population > 0) Then
    For run = population To 1 Step -1
        reaction.Remove (run)
    Next
End If
End Sub

Private Sub clear_braces()
Dim run As Integer
Dim population As Integer
population = brace.Count
If (population > 0) Then
    For run = population To 1 Step -1
        brace.Remove (run)
    Next
End If
End Sub
Private Sub Class_Terminate()
Set load = Nothing
Set reaction = Nothing
Set brace = Nothing
End Sub
