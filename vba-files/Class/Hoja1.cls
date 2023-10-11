VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
  
  Dim myRange As Range
  
  If Me.CodeName = "Hoja1" Then
    On Error Resume Next
    If Target.Column = 1 Then
      Set myRange = Range(Target, Target.Offset(, 3)).Select
    ElseIf Target.Column > 1 Then
      Set myRange = Target
    End If
    With Selection.Interior
      .Pattern = xlNone
      .TintAndShade = 0
      .PatternTintAndShade = 0
    End With
    On Error GoTo 0
  End If
End Sub

