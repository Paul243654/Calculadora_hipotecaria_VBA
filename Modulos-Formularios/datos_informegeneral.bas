Attribute VB_Name = "datos_informegeneral"
Option Explicit

Sub datos_informe_general()

Dim uftdp As Integer

Worksheets("dato_informe").Select
Worksheets("dato_informe").Cells.ClearContents

uftdp = Worksheets("cuadro_amortizacion").Cells(Rows.Count, 1).End(xlUp).Row

Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Range(Cells(1, 4), Cells(uftdp, 4)).Select
Selection.Copy
Worksheets("dato_informe").Select
Worksheets("dato_informe").Range(Cells(1, 1), Cells(uftdp, 1)).PasteSpecial xlPasteAll

Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Range(Cells(1, 6), Cells(uftdp, 6)).Select
Selection.Copy
Worksheets("dato_informe").Select
Worksheets("dato_informe").Range(Cells(1, 2), Cells(uftdp, 2)).PasteSpecial xlPasteAll

Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Range(Cells(1, 7), Cells(uftdp, 7)).Select
Selection.Copy
Worksheets("dato_informe").Select
Worksheets("dato_informe").Range(Cells(1, 3), Cells(uftdp, 3)).PasteSpecial xlPasteAll

Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Range(Cells(1, 11), Cells(uftdp, 11)).Select
Selection.Copy
Worksheets("dato_informe").Select
Worksheets("dato_informe").Range(Cells(1, 4), Cells(uftdp, 4)).PasteSpecial xlPasteAll

Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Range(Cells(1, 12), Cells(uftdp, 12)).Select
Selection.Copy
Worksheets("dato_informe").Select
Worksheets("dato_informe").Range(Cells(1, 5), Cells(uftdp, 5)).PasteSpecial xlPasteAll

Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Range(Cells(1, 16), Cells(uftdp, 16)).Select
Selection.Copy
Worksheets("dato_informe").Select
Worksheets("dato_informe").Range(Cells(1, 6), Cells(uftdp, 6)).PasteSpecial xlPasteAll

Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Range(Cells(1, 17), Cells(uftdp, 17)).Select
Selection.Copy
Worksheets("dato_informe").Select
Worksheets("dato_informe").Range(Cells(1, 7), Cells(uftdp, 7)).PasteSpecial xlPasteAll

Worksheets("dato_informe").Range("A1").Select

End Sub
