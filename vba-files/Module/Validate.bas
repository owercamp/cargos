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
  Dim num As Integer, x As Long
  Dim MyFile As Variant
  Dim Item, fso As Object
  Set fso = CreateObject("Scripting.FileSystemObject")

  On Error GoTo NotFound:
  ' Delete the existing SQL file
  fso.DeleteFile ("C:\Users\SOANDES-DSOFT\Documents\MACRO\testfile.sql")
  On Error GoTo 0

  ' Open the SQL file for appending
  Set MyFile = fso.OpenTextFile("C:\Users\SOANDES-DSOFT\Documents\MACRO\testfile.sql", ForAppending, True, TristateTrue)
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

  ThisWorkbook.ActiveSheet.Range(Selection.Offset(0, -3), Selection.Offset(0, -1)).Select
  Dim information As Variant, tblCargo As Object, tblCargoOrigin As Object
  information = Selection.Value

  ' Add the information to the tbl_cargo table
  Set tblCargoOrigin = Workbooks("Queries SQL SIGAD.xlsb").Worksheets("BASE P").ListObjects("tbl_cargo")

  For x = 1 To UBound(information, 1)
    With tblCargoOrigin.ListRows.Add
      .Range(1) = information(x, 1)
      .Range(2) = information(x, 2)
      .Range(3) = information(x, 3)
    End With
  DoEvents
  Next

  Range(Selection, Selection.Offset(, 1)).Select
  Selection.Style = "Notas"
  Range("A1").End(xlDown).Select
  ThisWorkbook.Save

  MsgBox "Importaci" & ChrW(243) & "n Completa", vbInformation + vbOKOnly, "Importaci" & ChrW(243) & "n SQL"

  ThisWorkbook.Close

  Exit Sub

NotFound:
  Resume Next
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

