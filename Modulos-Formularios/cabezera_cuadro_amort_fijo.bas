Attribute VB_Name = "cabezera_cuadro_amort_fijo"
Option Explicit

Sub cabezera_detalle_historial_fijo()

Worksheets("cuadro_amortizacion_fijo").Select
Worksheets("cuadro_amortizacion_fijo").Cells.ClearContents

Worksheets("cuadro_amortizacion_fijo").Range("A1").Value = "ncuota"
Worksheets("cuadro_amortizacion_fijo").Range("B1").Value = "cuota"
Worksheets("cuadro_amortizacion_fijo").Range("C1").Value = "interés"
Worksheets("cuadro_amortizacion_fijo").Range("D1").Value = "amortización"
Worksheets("cuadro_amortizacion_fijo").Range("E1").Value = "capital pendiente"

End Sub

