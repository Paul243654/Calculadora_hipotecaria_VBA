Attribute VB_Name = "detalle_cuotas_int_fijo"
Option Explicit

Sub detalle_cuotas_con_interes_fijo()

Dim capital_inicial As Double
Dim plazos As Integer
Dim amortizacion_cuota_fija As Double
Dim interes_cuota_fija As Double
Dim cuota_mes_fija As Double
Dim capital_pendiente_fijo As Double
Dim j_fijo As Double
Dim i_fijo As Double
Dim z As Integer
Dim uf As Integer
Dim n_cuota As Integer
Dim intereses_cuota_fija As Double
Dim total_intereses As Double

plazos = Worksheets("formulario_fijo").Range("B1").Value
capital_inicial = Worksheets("formulario_fijo").Range("B2").Value
interes_cuota_fija = Worksheets("formulario_fijo").Range("B3").Value
total_intereses = 0

capital_pendiente_fijo = capital_inicial
uf = Worksheets("cuadro_amortizacion_fijo").Range("A" & Rows.Count).End(xlUp).Row
               
    i_fijo = interes_cuota_fija / 1200
    j_fijo = (1 + i_fijo) ^ (-(plazos))
    cuota_mes_fija = (capital_pendiente_fijo * i_fijo) / (1 - j_fijo)
    
     For z = 1 To plazos
               intereses_cuota_fija = (capital_pendiente_fijo * interes_cuota_fija) / 1200
               amortizacion_cuota_fija = cuota_mes_fija - intereses_cuota_fija
               capital_pendiente_fijo = capital_pendiente_fijo - amortizacion_cuota_fija
               n_cuota = z
                                                                        
                                   Worksheets("cuadro_amortizacion_fijo").Select
                                   Worksheets("cuadro_amortizacion_fijo").Cells(uf + z, 1).Value = n_cuota
                                   Worksheets("cuadro_amortizacion_fijo").Cells(uf + z, 2).Value = cuota_mes_fija
                                   Worksheets("cuadro_amortizacion_fijo").Cells(uf + z, 3).Value = intereses_cuota_fija
                                   Worksheets("cuadro_amortizacion_fijo").Cells(uf + z, 4).Value = amortizacion_cuota_fija
                                   Worksheets("cuadro_amortizacion_fijo").Cells(uf + z, 5).Value = capital_pendiente_fijo
                                   
              total_intereses = total_intereses + intereses_cuota_fija
                                   
     Next z
     
     Worksheets("formulario_fijo").Range("B8").Value = cuota_mes_fija
     Worksheets("formulario_fijo").Range("B10").Value = total_intereses
     Worksheets("formulario_fijo").Range("B11").Value = ((100 / Worksheets("formulario_fijo").Range("B2").Value)) * total_intereses
     
End Sub


