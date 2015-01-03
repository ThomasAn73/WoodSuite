VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ld_counter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim load As Integer
Dim reaction As Integer
Dim brace As Integer

Public Sub load_inc()
load = load + 1
End Sub
Public Sub load_dec()
load = load - 1
'Debug.Print "LOad="; load
End Sub
Public Function load_count() As Integer
load_count = load
End Function

Public Sub reaction_inc()
reaction = reaction + 1
End Sub
Public Sub reaction_dec()
reaction = reaction - 1
End Sub
Public Function reaction_count() As Integer
reaction_count = reaction
End Function

Public Sub brace_inc()
brace = brace + 1
End Sub
Public Sub brace_dec()
brace = brace - 1
End Sub
Public Function brace_count() As Integer
brace_count = brace
End Function

Public Sub restart()
load = 0
reaction = 0
brace = 0
End Sub
