Attribute VB_Name = "parte_4"
Option Explicit
Public ufk As Integer

Sub resultados_finales()


ufk = Worksheets("cuadro_amortizacion").Range("A" & Rows.Count).End(xlUp).Row

Rem columnas que me interesan son : 4,6,7, 11, 12, 15, 16, 17

Worksheets("cuadro_amortizacion").Select

                                   Worksheets("cuadro_amortizacion").Cells(ufk + 1, 4).Value = Application.WorksheetFunction.CountA(Range(Cells(2, 4), Cells(ufk, 4)))
                                   Worksheets("cuadro_amortizacion").Cells(ufk + 1, 6).Value = Application.WorksheetFunction.Sum(Range(Cells(2, 6), Cells(ufk, 6)))
                                   Worksheets("cuadro_amortizacion").Cells(ufk + 1, 7).Value = Application.WorksheetFunction.Sum(Range(Cells(2, 7), Cells(ufk, 7)))
                                   Worksheets("cuadro_amortizacion").Cells(ufk + 1, 11).Value = Application.WorksheetFunction.Sum(Range(Cells(2, 11), Cells(ufk, 11)))
                                   Worksheets("cuadro_amortizacion").Cells(ufk + 1, 12).Value = Application.WorksheetFunction.Sum(Range(Cells(2, 12), Cells(ufk, 12)))
                                   Worksheets("cuadro_amortizacion").Cells(ufk + 1, 15).Value = Application.WorksheetFunction.Sum(Range(Cells(2, 15), Cells(ufk, 15)))
                                   Worksheets("cuadro_amortizacion").Cells(ufk + 1, 16).Value = Application.WorksheetFunction.Sum(Range(Cells(2, 16), Cells(ufk, 16)))
                                   Worksheets("cuadro_amortizacion").Cells(ufk + 1, 17).Value = Application.WorksheetFunction.Sum(Range(Cells(2, 17), Cells(ufk, 17)))


End Sub
