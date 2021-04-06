Option Explicit

Sub actualizar_datos()
'Abre la planilla de hoy y arma el reporte con los datos con los valores actualizados.
'En caso de encontrarse con que la planilla más actualizada está en blanco por algún problema de que no la cargaron en Catedral
'busca los datos de la planilla anterior. Ya si hay 2 días seguidos de planillas en blanco, deja de funcionar la macro. Chequea 3 nombres posibles para la planilla
'por si le tuvieron que cambiar el nombre a Copia de ** o a ** - copia. Pero si existen dos archivos (normal y "Copia de") utiliza siempre el normal
'Esto último puede ser un problema siendo que el archivo "Copia de" se puede crear sin borrar el normal.

Dim sDireccionPlanilla As String
Dim wPlanillaDiaria As Workbook
Dim wReporte As Workbook
Dim sHoy As String
Dim sAyer As String
Dim lEquipos As Long


'Abre la planilla diaria de HOY y la define como objeto
Set wReporte = ThisWorkbook
wReporte.Worksheets("FECHA").Range("C2").Value = Format(Date, "mmmm")  'Establece mes de HOY en planilla Reporte
wReporte.Worksheets("FECHA").Range("D2").Value = Format(Date, "yyyy")  'Establece año de HOY en planilla Reporte
sDireccionPlanilla = Worksheets("FECHA").Range("F4").Value

If Dir(sDireccionPlanilla) = "" Then 'Chequeos de otros nombres posibles, en nested if para establecer un orden de prioridad para su apertura
        sDireccionPlanilla = "\\mtvdfs01\DireccionOperativa\Compart_FS01\Compartido\SS.AA. Catedral\Reporte de escaleras y ascensores\" _
        & Year(Date) & "\Copia de Reporte de Escaleras y Elevadores - " & Format(Date, "mmmm") & " " & Year(Date) & ".xls"
        If Dir(sDireccionPlanilla) = "" Then
            sDireccionPlanilla = "\\mtvdfs01\DireccionOperativa\Compart_FS01\Compartido\SS.AA. Catedral\Reporte de escaleras y ascensores\" _
            & Year(Date) & "\Reporte de Escaleras y Elevadores - " & Format(Date, "mmmm") & " " & Year(Date) & " - copia.xls"
            wReporte.Worksheets("Recopilador").Range("B2:Z97").Replace ".xls", " - copia.xls" 'redirige el recopilador al nuevo nombre
        Else
        wReporte.Worksheets("Recopilador").Range("B2:Z97").Replace "Reporte de Escaleras y Elevadores - ", "Copia de Reporte de Escaleras y Elevadores - " 'redirige el recopilador al nuevo nombre
        End If
End If

Workbooks.Open sDireccionPlanilla, ReadOnly:=True
Set wPlanillaDiaria = ActiveWorkbook

'Chequea si la planilla de hoy ya tiene aunque
'sea 10 equipos cargados, caso contrario asume
'que falta actualizar y toma los valores de ayer a última hora.
'Para el caso de que haya 0 equipos, el método Count colapsa a menos que lEquipos sea Long, y se englobe en el on error
sHoy = Format(Now(), "dd")
sAyer = Format(Date - 1, "dd")

On Error Resume Next
lEquipos = wPlanillaDiaria.Worksheets(sHoy).Range("C5:C100").Cells.SpecialCells(xlCellTypeConstants).Count
On Error GoTo 0

If lEquipos > 10 Then
    ThisWorkbook.Worksheets("FECHA").Range("A2").Value = sHoy
Else
    If sHoy = 1 Then 'Justo cayó cambio de mes, así que corresponde redefinir wPlanillaDiaria por la del mes pasado para tomar el último día del mismo
        wPlanillaDiaria.Close (False)
        wReporte.Worksheets("FECHA").Range("C2").Value = Format(Date - 1, "mmmm")
        wReporte.Worksheets("FECHA").Range("D2").Value = Format(Date - 1, "yyyy")
        sDireccionPlanilla = Worksheets("FECHA").Range("F4").Value
        
        If Dir(sDireccionPlanilla) = "" Then 'Chequeos de otros nombres posibles, en nested if para establecer un orden de prioridad para su apertura
                sDireccionPlanilla = "\\mtvdfs01\DireccionOperativa\Compart_FS01\Compartido\SS.AA. Catedral\Reporte de escaleras y ascensores\" _
                & Year(Date - 1) & "\Copia de Reporte de Escaleras y Elevadores - " & Format(Date - 1, "mmmm") & " " & Year(Date - 1) & ".xls"
                If Dir(sDireccionPlanilla) = "" Then
                    sDireccionPlanilla = "\\mtvdfs01\DireccionOperativa\Compart_FS01\Compartido\SS.AA. Catedral\Reporte de escaleras y ascensores\" _
                    & Year(Date - 1) & "\Reporte de Escaleras y Elevadores - " & Format(Date - 1, "mmmm") & " " & Year(Date - 1) & " - copia.xls"
                End If
        End If
             
        
        Workbooks.Open sDireccionPlanilla, ReadOnly:=True
        Set wPlanillaDiaria = ActiveWorkbook
        wReporte.Worksheets("FECHA").Range("A2").Value = sAyer
    Else
        ThisWorkbook.Worksheets("FECHA").Range("A2").Value = sAyer
    End If
End If

'Actualiza todo y activa la pestaña Reporte
ThisWorkbook.RefreshAll
ThisWorkbook.Worksheets("Reporte").Activate

'Cierra la Planilla Diaria y vacía las variables
wPlanillaDiaria.Close (False)

Set wPlanillaDiaria = Nothing
Set wReporte = Nothing


End Sub