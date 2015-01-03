VERSION 5.00
Begin VB.Form frm_graph 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Graph"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   8505
   Icon            =   "frm_graph.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   8505
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TextShear 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   1575
   End
   Begin VB.PictureBox TheShear 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   240
      ScaleHeight     =   2715
      ScaleWidth      =   7995
      TabIndex        =   5
      ToolTipText     =   "Click to evoke the Bending Diagram."
      Top             =   600
      Width           =   8055
   End
   Begin VB.TextBox TextMoment 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6720
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Top             =   3480
      Width           =   1095
   End
   Begin VB.PictureBox picBox 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   2775
      Left            =   240
      ScaleHeight     =   2715
      ScaleWidth      =   7995
      TabIndex        =   0
      ToolTipText     =   "Click to evoke the Shear Diagram."
      Top             =   600
      Width           =   8055
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Shear Diagram"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Maximum Shear"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Maximum Moment"
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Bending Moment Diagram"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frm_graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frm_graph.Hide
End Sub

Public Sub draw(DrawHere As PictureBox, x As Integer, y As Double, Optional xbound As Double, Optional ybound As Double, Optional thiscolor As Long)


DrawHere.Scale (0, ybound + 1)-(xbound, 0)

DrawHere.Line (x, ybound / 2)-(x + 1, y + ybound / 2), thiscolor, B

End Sub

Private Sub Form_Activate()
Dim x As Long
Dim y As Long

Call bendingView(True)
Call shearView(False)
y = picBox.Height
x = picBox.width
picBox.Scale (0, y)-(x, 0)
'Draw a datum line
picBox.Line (0, y / 2)-(x, y / 2)

End Sub

Private Sub bendingView(status As Boolean)
picBox.Visible = status
Label1.Visible = status
Label2.Visible = status
TextMoment.Visible = status
End Sub

Private Sub shearView(status As Boolean)
TheShear.Visible = status
Label4.Visible = status
Label3.Visible = status
TextShear.Visible = status
End Sub


Private Sub picBox_Click()
Call shearView(True)
Call bendingView(False)
End Sub

Private Sub TheShear_Click()
Call shearView(False)
Call bendingView(True)
End Sub
