Attribute VB_Name = "Validate"
Option Explicit

Public Sub Checked()
  ' This subroutine sets the style of the Selection to "Notas"
  ' It is invoked when the Checked attribute is triggered
  Selection.Style = "Notas"
End Sub

Public Sub btn_SQL_Click()
  ' This sub is executed when the SQL button is clicked

  Dim data As Range
  Dim num As Integer
  Dim MyFile As Variant
  Dim Item, fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")
  
  Dim btn As MsoButtonState
  
  btn = MsgBox(ChrW(191) + "desea anexar nueva informaci" + ChrW(243) + "n al archivo testfile.sql" + ChrW(63), vbDefaultButton1 + vbExclamation + vbYesNo, "SQL Cargos")
  
  If btn = vbNo Then
    ' Delete the existing SQL file
    If fso.FileExists(Workbooks("Queries SQL SIGAD.xlsb").Worksheets("RUTAS").Range("$C$9").Value & "\testfile.sql") Then
      fso.DeleteFile (Workbooks("Queries SQL SIGAD.xlsb").Worksheets("RUTAS").Range("$C$9").Value & "\testfile.sql")
    End If
  End If

  ' Open the SQL file for appending
  Set MyFile = fso.OpenTextFile(Workbooks("Queries SQL SIGAD.xlsb").Worksheets("RUTAS").Range("$C$9").Value & "\testfile.sql", ForAppending, True, TristateTrue)
  Set data = Selection
  num = data.CountLarge

  ' Write the initial SQL statement
  MyFile.WriteLine "INSERT INTO cargos (`id`,`id_categoria_cargo`,`nombre`) VALUES"
  For Each Item In data
    If Item <> "" And num <> 1 Then
      ' Write the SQL line for each item
      MyFile.WriteLine Item
      num = num - 1
    ElseIf Item <> "" And num = 1 Then
      ' Write the last SQL line with the last item
      MyFile.WriteLine reemplazarUltimoCaracter(Item)
      num = num - 1
    End If
  Next Item
  MyFile.WriteLine ""
  MyFile.Close

  ActiveCell.Offset(0, -3).Select
  Dim information As Object, book As String
  book = ThisWorkbook.Name
  If ActiveCell.Offset(1, 0).Value <> vbNullString Then
    Set information = Range(ActiveCell, ActiveCell.End(xlDown))
  Else
    Set information = Range(ActiveCell, ActiveCell)
  End If

  With Application
    .ScreenUpdating = False
    .EnableEvents = False
    .Calculation = xlCalculationManual
  End With
  
  ' Add the information to the tbl_cargo table
  Windows("Queries SQL SIGAD.xlsb").Activate
  Worksheets("BASE P").Select
  Range("tbl_cargo").End(xlDown).Offset(1, 0).Select

  num = 0
  For Each Item In information
    ActiveCell = Item.Value
    ActiveCell.Offset(0, 1) = Item.Offset(, 1).Value
    ActiveCell.Offset(0, 2) = Item.Offset(, 2).Value
    num = num + 1
    Application.StatusBar = "Importando: " & CStr(num)
    ActiveCell.Offset(1, 0).Select
  DoEvents
  Next Item

  Worksheets("TRABAJADORES").Select
  Windows(book).Activate
  information.Select
  Range(Selection, Selection.Offset(, 3)).Select
  Selection.Style = "Notas"
  Range("A1").End(xlDown).Select

  With Application
    .ScreenUpdating = True
    .EnableEvents = True
    .Calculation = xlCalculationAutomatic
  End With

  ThisWorkbook.Save

  MsgBox "Importaci" & ChrW(243) & "n Completa", vbInformation + vbOKOnly, "Importaci" & ChrW(243) & "n SQL"
  
  ThisWorkbook.Close

End Sub

Private Function reemplazarUltimoCaracter(ByVal texto As String) As String
  ' This function replaces the last character of the given text with a semicolon (;)

  Dim ultimoCaracter As String
  ultimoCaracter = ","

  Dim posicion As Integer
  posicion = InStrRev(texto, ultimoCaracter)

  If posicion > 0 Then
    ' If the last character is found in the text
    ' Replace the last character with a semicolon (;)
    reemplazarUltimoCaracter = Left(texto, posicion - 1) & ";" & Right(texto, Len(texto) - posicion)
  End If

  ' Return the modified text
End Function

