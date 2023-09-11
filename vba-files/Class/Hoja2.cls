VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
  On Error Resume Next
  If Target.Column = 1 And Me.CodeName = "Hoja2" Then
    Application.EnableEvents = False
    Application.ScreenUpdating = False

    Dim rngData As Range
    Set rngData = Target

    With rngData
        .RemoveDuplicates Columns:=1, Header:=xlNo
    End With
    
    ActiveWorkbook.Worksheets("DUPLICADOS").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("DUPLICADOS").Sort.SortFields.Add Key:=Range("$A1") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("DUPLICADOS").Sort
        .SetRange Range("$A1:$A10000")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
  End If
End Sub

