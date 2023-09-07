Attribute VB_Name = "Validate"
Option Explicit

Public Sub Checked()
Attribute Checked.VB_ProcData.VB_Invoke_Func = "j\n14"
    Selection.Style = "Notas"
End Sub

Public Sub btn_SQL_Click()
    Dim data As Range
    Dim num As Integer, x As Long
    Dim MyFile As Variant
    Dim Item, FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
        
    On Error GoTo NotFound:
    FSO.DeleteFile ("C:\Users\SOANDES-DSOFT\Documents\MACRO\testfile.sql")
    On Error GoTo 0

    Set MyFile = FSO.OpenTextFile("C:\Users\SOANDES-DSOFT\Documents\MACRO\testfile.sql", ForAppending, True, TristateTrue)
    Set data = Selection
    num = data.CountLarge

    MyFile.WriteLine "INSERT INTO cargos (`id`,`id_categoria_cargo`,`nombre`) VALUES"
    For Each Item In data
        If Item <> "" And num <> 1 Then
            MyFile.WriteLine Item
            num = num - 1
        ElseIf Item <> "" And num = 1 Then
            MyFile.WriteLine reemplazarUltimoCaracter(Item)
            num = num - 1
        End If
    Next Item
    MyFile.WriteLine ""
    MyFile.Close

    Range(Selection.Offset(, -3), Selection.Offset(, -1)).Select
    Dim information As Variant, tblCargo As Object, tblCargoOrigin As Object
    information = Selection.Value

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
    
    MsgBox "Importación Completa", vbInformation + vbOKOnly, "Importación SQL"
    
    ThisWorkbook.Close
    
    Exit Sub
    
NotFound:
    Resume Next
End Sub

Private Function reemplazarUltimoCaracter(ByVal texto As String) As String
  Dim ultimoCaracter As String

  ultimoCaracter = ","

  Dim posicion As Integer
  posicion = InStrRev(texto, ultimoCaracter)

  If posicion > 0 Then
    reemplazarUltimoCaracter = Left(texto, posicion - 1) & ";" & Right(texto, Len(texto) - posicion)
  End If
End Function

