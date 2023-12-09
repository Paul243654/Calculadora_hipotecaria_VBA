Attribute VB_Name = "datos_para_tabla"
Option Explicit

Sub copiar_datos_para_tabla()

Dim uftdp As Integer

Worksheets("datos_tabla").Select
Worksheets("datos_tabla").Cells.ClearContents

uftdp = Worksheets("cuadro_amortizacion").Cells(Rows.Count, 1).End(xlUp).Row

Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Range(Cells(1, 1), Cells(uftdp, 1)).Select
Selection.Copy
Worksheets("datos_tabla").Select
Worksheets("datos_tabla").Range(Cells(1, 1), Cells(uftdp, 1)).PasteSpecial xlPasteAll

Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Range(Cells(1, 4), Cells(uftdp, 4)).Select
Selection.Copy
Worksheets("datos_tabla").Select
Worksheets("datos_tabla").Range(Cells(1, 2), Cells(uftdp, 2)).PasteSpecial xlPasteAll

Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Range(Cells(1, 5), Cells(uftdp, 5)).Select
Selection.Copy
Worksheets("datos_tabla").Select
Worksheets("datos_tabla").Range(Cells(1, 3), Cells(uftdp, 3)).PasteSpecial xlPasteAll

Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Range(Cells(1, 10), Cells(uftdp, 10)).Select
Selection.Copy
Worksheets("datos_tabla").Select
Worksheets("datos_tabla").Range(Cells(1, 4), Cells(uftdp, 4)).PasteSpecial xlPasteAll

Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Range(Cells(1, 9), Cells(uftdp, 9)).Select
Selection.Copy
Worksheets("datos_tabla").Select
Worksheets("datos_tabla").Range(Cells(1, 5), Cells(uftdp, 5)).PasteSpecial xlPasteAll

Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Range(Cells(1, 14), Cells(uftdp, 14)).Select
Selection.Copy
Worksheets("datos_tabla").Select
Worksheets("datos_tabla").Range(Cells(1, 6), Cells(uftdp, 6)).PasteSpecial xlPasteAll

Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Range(Cells(1, 15), Cells(uftdp, 15)).Select
Selection.Copy
Worksheets("datos_tabla").Select
Worksheets("datos_tabla").Range(Cells(1, 7), Cells(uftdp, 7)).PasteSpecial xlPasteAll

Worksheets("datos_tabla").Range("A1").Select


End Sub

