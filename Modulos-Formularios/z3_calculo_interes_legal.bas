Attribute VB_Name = "z3_calculo_interes_legal"
Option Explicit

Sub cobro_interes_LLL()

Rem #5

Dim fecha_inicial As Date
Dim fecha_final As Date
Dim Parametro As String
Dim monto As Double
Dim i As Integer
Dim j As Integer
Dim interes_legal As Double
Dim dias As Integer
Dim intereses As Double
Dim n_cuota As Integer
Dim uf As Integer
Dim ufv As Integer
Dim tabla As Range
Dim tabla_interes As Range
Dim buscar_numero_fila As Double
Dim referencia As Integer
Dim f As Double
Dim año_cuota As Integer
Dim año_interes As Integer
Dim ufdatos As Integer
Dim fecha_leg As Date

Dim obj_Cell As Range
Dim comodin As Integer
Dim fila_buscada As Integer
Dim val_enc As Integer
    
comodin = Worksheets("formulario").Range("B7").Value

Worksheets("datos").Range("I2:I30").Select
fila_buscada = 1
val_enc = 30

    For Each obj_Cell In Selection.Cells
          
       If obj_Cell.Value = comodin Then
            fila_buscada = fila_buscada + 1
            val_enc = fila_buscada
       Else:
             fila_buscada = fila_buscada + 1
       End If

    Next

Parametro = "d"

Worksheets("datos_volcados").Range("A2:I15000").ClearContents
Worksheets("datos_volcados").Range("A1").Value = "Añocobro"
Worksheets("datos_volcados").Range("B1").Value = "ncuota"
Worksheets("datos_volcados").Range("C1").Value = "Cobradodemas (€)"
Worksheets("datos_volcados").Range("D1").Value = "Fechainicial"
Worksheets("datos_volcados").Range("E1").Value = "fechafinal"
Worksheets("datos_volcados").Range("F1").Value = "ndias"
Worksheets("datos_volcados").Range("G1").Value = "Interéslegaldeldinero"
Worksheets("datos_volcados").Range("H1").Value = "InteresLegal"


fecha_leg = Format(CDate(InputBox("Ingresar fecha final para cálculo de intereses:  dd - mm - aaaa ")), "dd-mm-yyyy")


Worksheets("datos").Cells(val_enc, 14) = fecha_leg

Rem buscar ultima fila y añadirle uno (ufv)
ufdatos = Worksheets("datos").Cells(Rows.Count, 1).End(xlUp).Row


For i = 2 To ufdatos
     
     fecha_inicial = Worksheets("datos").Cells(i, 3)
     monto = Worksheets("datos").Cells(i, 4)
     año_cuota = Worksheets("datos").Cells(i, 5)

     Set tabla = Worksheets("datos").Range("I2:N30")
     referencia = Worksheets("datos").Cells(i, 1)
     buscar_numero_fila = Application.VLookup(referencia, tabla, 4, False)
     f = buscar_numero_fila
     

          For j = f To val_enc

     
               interes_legal = Worksheets("datos").Cells(j, 10)
               fecha_final = Worksheets("datos").Cells(j, 14)
               dias = DateDiff(Parametro, fecha_inicial, fecha_final)
               intereses = (dias / 365) * (monto) * (interes_legal)
               n_cuota = i - 1
               
               Rem buscar ultima fila y añadirle uno (ufv)
               uf = Worksheets("datos_volcados").Cells(Rows.Count, 1).End(xlUp).Row
               ufv = uf + 1
               
               Worksheets("datos_volcados").Cells(ufv, 1) = año_cuota
               Worksheets("datos_volcados").Cells(ufv, 2) = n_cuota
               Worksheets("datos_volcados").Cells(ufv, 3) = monto
               Worksheets("datos_volcados").Cells(ufv, 4) = fecha_inicial
               Worksheets("datos_volcados").Cells(ufv, 5) = fecha_final
               Worksheets("datos_volcados").Cells(ufv, 6) = dias
               Worksheets("datos_volcados").Cells(ufv, 7) = interes_legal
               Worksheets("datos_volcados").Cells(ufv, 8) = intereses
               
               fecha_inicial = Worksheets("datos").Cells(j + 1, 13)
               

          Next j


Next i

Worksheets("datos_volcados").Select Range("A2").Select
Worksheets("datos_volcados").Columns("A:H").AutoFit

End Sub



