Attribute VB_Name = "z1_cop_peg_informacion_tratada"

Sub copiaypega_informacion_decalculadora()

Dim ufaEliminar As Integer
Dim ufi As Integer

Rem #1

Worksheets("datos_iniciales").Select
Worksheets("datos_iniciales").Range("A1:Z1000").Clear

Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Range("A:Q").Columns.Copy
Worksheets("datos_iniciales").Select
Worksheets("datos_iniciales").Range("A:Q").Columns.PasteSpecial

ufaEliminar = Worksheets("datos_iniciales").Cells(Rows.Count, 1).End(xlUp).Row
ufi = ufaEliminar + 1

Worksheets("datos_iniciales").Range(Cells(ufi, 1), Cells(ufi, 17)).Clear

Worksheets("datos_iniciales").Columns("A:Q").AutoFit

End Sub





