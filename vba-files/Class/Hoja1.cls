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
  
  If Me.CodeName = "Hoja1" And Target.Column = 1 Or Target.Column = 4 Then
    Range(Target, Target.Offset(, 3)).Select
    With Selection.Interior
      .Pattern = xlNone
      .TintAndShade = 0
      .PatternTintAndShade = 0
    End With
  ElseIf Target.Column < 3 Then
    Target.Select
    With Selection.Interior
      .Pattern = xlNone
      .TintAndShade = 0
      .PatternTintAndShade = 0
    End With
  End If
End Sub

