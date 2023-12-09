Attribute VB_Name = "z6_funcion_coincidir_index"
Sub ult_dd()

Rem #8

Dim ufaEliminar As Integer
Dim ufi As Integer


Worksheets("td_transitoria").Select
Worksheets("td_transitoria").Range("A1:C1000").Clear

Worksheets("tabla_int_legal").Select
Worksheets("tabla_int_legal").Range("A1:C1000").Copy
Worksheets("td_transitoria").Select
Worksheets("td_transitoria").Range("A1:C1000").PasteSpecial xlValues

ufi = Worksheets("td_transitoria").Cells(Rows.Count, 1).End(xlUp).Row

Worksheets("td_transitoria").Range(Cells(ufi, 1), Cells(ufi, 3)).Clear

Worksheets("datos_iniciales").Columns("A:C").AutoFit

End Sub
 

Sub coincidir_indice()

Rem #9

Dim lfila As Integer
Dim i As Integer
Dim j As Integer
Dim tablet As Range
Dim cuota As Integer
Dim buscarv_interes_legal As Variant
Dim interes_legal_acumulado As Variant
Dim valor As Variant


lfila = Worksheets("datos_con_int_legal").Cells(Rows.Count, 1).End(xlUp).Row
Worksheets("datos_con_int_legal").Cells(1, 18).Value = "Interés Legal"

Set tablet = Worksheets("td_transitoria").Range("A2:C1000")

    For i = 2 To lfila
    
    cuota = Worksheets("datos_con_int_legal").Cells(i, 4).Value
    buscarv_interes_legal = Application.VLookup(cuota, tablet, 3, 0)
    Worksheets("datos_con_int_legal").Cells(i, 18).Value = buscarv_interes_legal
    interes_legal_acumulado = interes_legal_acumulado + buscarv_interes_legal
    
    Next i

    For j = 2 To lfila
    
        Worksheets("datos_con_int_legal").Cells(j, 18).NumberFormat = "#,##0.00"
        Worksheets("datos_con_int_legal").Cells(j, 18).Font.Size = 9
        Worksheets("datos_con_int_legal").Cells(j, 18).HorizontalAlignment = xlCenter
        
    Next j
    
valor = Round(interes_legal_acumulado, 2)
MsgBox "El interés total acumulado de todas las cuotas asciende a: " & valor

End Sub
