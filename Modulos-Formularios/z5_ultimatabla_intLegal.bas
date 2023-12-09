Attribute VB_Name = "z5_ultimatabla_intLegal"
Sub copiaypega_informacion_decalculadora_conintereslegal()

Rem #7

Dim ufaEliminar As Integer
Dim ufi As Integer


Worksheets("datos_con_int_legal").Select
Worksheets("datos_con_int_legal").Range("A1:Z1000").Clear

Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Range("A:Q").Columns.Copy
Worksheets("datos_con_int_legal").Select
Worksheets("datos_con_int_legal").Range("A:Q").Columns.PasteSpecial

ufaEliminar = Worksheets("datos_con_int_legal").Cells(Rows.Count, 1).End(xlUp).Row
ufi = ufaEliminar + 1

Worksheets("datos_con_int_legal").Range(Cells(ufi, 1), Cells(ufi, 17)).Clear

Worksheets("datos_con_int_legal").Columns("A:Q").AutoFit

End Sub
