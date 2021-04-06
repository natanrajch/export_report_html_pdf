Function ArrayToTextFile(a As Variant, strPath As String)
Dim fso         As Object
Dim fn          As Object
Dim i           As Long

Set fso = CreateObject("Scripting.FileSystemObject")
Set fn = fso.OpenTextFile(strPath, 2, True)

'For i = LBound(a) To UBound(a)
    fn.writeline a '(i)
'Next i

fn.Close
End Function

Sub CrearMapaHTML()

Dim a As Variant
Dim i As Integer
Dim colDIV As Integer
Dim colSTYLE As Integer
Dim wb As Workbook
Dim strPath As String

Set wb = ThisWorkbook

'Crea primeras líneas del HTML
a = wb.Worksheets("PropiedadesHTML").Range("HTML1").Value & vbCrLf
a = a & wb.Worksheets("PropiedadesHTML").Range("HTML2").Value & vbCrLf
a = a & wb.Worksheets("PropiedadesHTML").Range("HTML3").Value & vbCrLf
a = a & wb.Worksheets("PropiedadesHTML").Range("HTML4").Value & vbCrLf
a = a & wb.Worksheets("PropiedadesHTML").Range("HTML5").Value & vbCrLf

'Inserta los semáforos de estado
colDIV = wb.Worksheets("UbicacionEstaciones").Range("PrimerDIV").Column
For i = wb.Worksheets("UbicacionEstaciones").Range("PrimerDIV").Row To wb.Worksheets("UbicacionEstaciones").Range("PrimerDIV").End(xlDown).Row
    a = a & wb.Worksheets("UbicacionEstaciones").Cells(i, colDIV).Value & vbCrLf
Next i

'Inserta siguientes líneas HTML
a = a & wb.Worksheets("PropiedadesHTML").Range("HTML5bis").Value
a = a & wb.Worksheets("PropiedadesHTML").Range("HTML5ter").Value
a = a & wb.Worksheets("PropiedadesHTML").Range("HTML5cuar").Value
a = a & wb.Worksheets("PropiedadesHTML").Range("HTML6").Value


'Inserta las propiedades de los semáforos de estado
colSTYLE = wb.Worksheets("UbicacionEstaciones").Range("PrimerSTYLE").Column
For i = wb.Worksheets("UbicacionEstaciones").Range("PrimerSTYLE").Row To wb.Worksheets("UbicacionEstaciones").Range("PrimerSTYLE").End(xlDown).Row
    a = a & wb.Worksheets("UbicacionEstaciones").Cells(i, colSTYLE).Value & vbCrLf
Next i

'Inserta últimas líneas HTML
a = a & wb.Worksheets("PropiedadesHTML").Range("HTML7").Value

'De necesitar cambiar la ubicación donde se crea el HTML, cambiar la siguiente línea:
strPath = wb.Path & "\MapaHTML-" & Format((Date), "dd-mm-yy") & ".html"

'Crea el archivo HTML con el mapa:
ArrayToTextFile a, strPath


Set wb = Nothing
End Sub

Sub BotónMapa()
Call CrearMapaHTML
MsgBox ("Mapa creado en: " & vbNewLine & vbNewLine & ThisWorkbook.Path & vbNewLine & vbNewLine _
        & "Con el nombre de archivo: " & vbNewLine & "MapaHTML-" & Format((Date), "dd-mm-yy") & ".html")

End Sub