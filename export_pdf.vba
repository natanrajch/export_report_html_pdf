Option Explicit

Sub ExportarReporteAPDF() 'espero les sirva!
Dim r As Integer
Dim rReporte As Range
r = 8

'Se fija hasta qué fila llega la tabla. xlEnd no funciona porque las celdas con fórmulas se extienden más allá
While ThisWorkbook.Worksheets("Reporte").Cells(r, 7).Value <> ""
    r = r + 1
Wend

With ThisWorkbook.Worksheets("Reporte")
Set rReporte = .Range(.Cells(1, 1), .Cells(r, 8))
End With

'Acá abajo actualizar dirección y nombre del archivo
rReporte.ExportAsFixedFormat Type:=xlTypePDF, Filename:= _
    "C:\ReporteMediosElevación.pdf", Quality:= _
    xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
    OpenAfterPublish:=True

End Sub

