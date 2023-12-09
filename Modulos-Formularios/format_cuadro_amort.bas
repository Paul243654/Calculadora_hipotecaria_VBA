Attribute VB_Name = "format_cuadro_amort"
Option Explicit

Sub format_cuadro_amortizacion()

Worksheets("cuadro_amortizacion").Select

With Worksheets("cuadro_amortizacion").Range("A1:S480")
 .Font.Size = 9
End With

With Worksheets("informe").Range("E2:H480")
 .NumberFormat = "#,##0.00"
End With

With Worksheets("informe").Range("I2:I480")
 .NumberFormat = "#,##0.000"
End With

With Worksheets("informe").Range("J2:M480")
 .NumberFormat = "#,##0.00"
End With

With Worksheets("informe").Range("N2:N480")
 .NumberFormat = "#,##0.000"
End With

With Worksheets("informe").Range("O2:Q480")
 .NumberFormat = "#,##0.00"
End With

Range("A1:S480").Select
Cells.Select
Cells.EntireColumn.AutoFit
Cells.EntireRow.AutoFit

Range("A2").Select

End Sub
