Attribute VB_Name = "z2_volcado_datos"
Sub Elimina_Columnas_int_legal()

Rem #2

Dim ufdatos As Integer
Dim Fecha As Date

Worksheets("datos_iniciales").Select
Worksheets("datos_iniciales").Range("E:P").Columns.Delete
Worksheets("datos_iniciales").Range("A:A").Columns.Delete
Worksheets("datos_iniciales").Range("C:C").Columns.Insert
Worksheets("datos_iniciales").Range("C1").Value = "Fecha de pago"
Worksheets("datos_iniciales").Range("D:D").Columns.Copy
Worksheets("datos_iniciales").Range("F:F").Columns.PasteSpecial
Worksheets("datos_iniciales").Range("D:D").Columns.Delete
Worksheets("datos_iniciales").Range("E:E").Columns.Insert
Worksheets("datos_iniciales").Range("E1").Value = "Año de cuota"
Worksheets("datos_iniciales").Range("A1").Value = "Año H"
Worksheets("datos_iniciales").Range("B:B").Columns.Insert
Worksheets("datos_iniciales").Range("B1").Value = "Año"


ufdatos = Worksheets("datos_iniciales").Cells(Rows.Count, 1).End(xlUp).Row

    
    For i = 2 To ufdatos
    
        If (Worksheets("datos_iniciales").Cells(i, 1).Value) = 2009 And (Worksheets("datos_iniciales").Cells(i, 3).Value) > 3 And (Worksheets("datos_iniciales").Cells(i, 3).Value) < 13 Then
        
            Worksheets("datos_iniciales").Cells(i, 2).Value = 20091
        
        Else

            Worksheets("datos_iniciales").Cells(i, 2).Value = Worksheets("datos_iniciales").Cells(i, 1).Value
            
        End If
        
    Next i
    

Worksheets("datos_iniciales").Range("A:A").Columns.Copy
Worksheets("datos_iniciales").Range("F:F").Columns.PasteSpecial
Worksheets("datos_iniciales").Range("A:A").Columns.Delete
Worksheets("datos_iniciales").Range("E1").Value = "Año de cuota"

Rem Worksheets("datos_iniciales").Range("C:C").Select
Rem Selection.NumberFormat = "dd/mm/yyyy"


End Sub


'Sub fechador()
'
'Dim FirstDate As Date
'Dim IntervalType As String
'Dim Number As Integer
'Dim Msg As String
'IntervalType = "m"    ' "m" specifies months as interval.
'FirstDate = InputBox("Enter a date")
'Number = InputBox("Enter number of months to add")
'Msg = "New date: " & DateAdd(IntervalType, Number, FirstDate)
'MsgBox Msg
'
'End Sub

Sub Ingresar_Fecha()

Rem #3

Dim Fecha As Date
Dim FirstDate As Date
Dim IntervalType As String
Dim Number As Integer
Dim Fecha_Devuelta As String
Dim uf As Integer

Fecha = Format(CDate(InputBox("1ra fecha de pago en formato:  dd - mm - aaaa ")), "dd-mm-yyyy")
Worksheets("datos_iniciales").Range("C2") = Fecha

uf = Worksheets("datos_iniciales").Cells(Rows.Count, 1).End(xlUp).Row

IntervalType = "m"    ' "m" especifica mes como intervalo


Number = 1

    For j = 3 To uf
    
        Fecha_Devuelta = Format(CDate(DateAdd(IntervalType, Number, Fecha)), "mm-dd-yyyy")
        Worksheets("datos_iniciales").Cells(j, 3) = Fecha_Devuelta
        Fecha = Format(CDate(Fecha_Devuelta), "mm-dd,yyyy")
        
    Next j
    
Worksheets("datos_iniciales").Columns("A:I").AutoFit


End Sub

Sub CopiayPega_informacion_tt()

Rem #4

Worksheets("datos").Select
Worksheets("datos").Range("A2:F600").ClearContents
Worksheets("datos_iniciales").Select
Worksheets("datos_iniciales").Range("A:F").Columns.Copy
Worksheets("datos").Select
Worksheets("datos").Range("A:F").Columns.PasteSpecial
Worksheets("datos").Columns("A:F").AutoFit

End Sub





