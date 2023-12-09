Attribute VB_Name = "parte_0"
Option Explicit

Sub parte_0_borrado_datos()


Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Cells.ClearContents

Worksheets("cuadro_amortizacion").Range("A1").Value = "nrev"
Worksheets("cuadro_amortizacion").Range("B1").Value = "año"
Worksheets("cuadro_amortizacion").Range("C1").Value = "mes"
Worksheets("cuadro_amortizacion").Range("D1").Value = "ncuota"
Worksheets("cuadro_amortizacion").Range("E1").Value = "cuota_irph"
Worksheets("cuadro_amortizacion").Range("F1").Value = "int_irph"
Worksheets("cuadro_amortizacion").Range("G1").Value = "amort_irph"
Worksheets("cuadro_amortizacion").Range("H1").Value = "cap pte_irph"
Worksheets("cuadro_amortizacion").Range("I1").Value = "irph"
Worksheets("cuadro_amortizacion").Range("J1").Value = "cuota_eur"
Worksheets("cuadro_amortizacion").Range("K1").Value = "int_eur"
Worksheets("cuadro_amortizacion").Range("L1").Value = "amort_eur"
Worksheets("cuadro_amortizacion").Range("M1").Value = "cap pte_eur"
Worksheets("cuadro_amortizacion").Range("N1").Value = "euribor"
Worksheets("cuadro_amortizacion").Range("O1").Value = "dif_cuotas"
Worksheets("cuadro_amortizacion").Range("P1").Value = "a_amort"
Worksheets("cuadro_amortizacion").Range("Q1").Value = "devolver"
Worksheets("cuadro_amortizacion").Range("R1").Value = "coincidir"
Worksheets("cuadro_amortizacion").Range("S1").Value = "% progreso"

End Sub
