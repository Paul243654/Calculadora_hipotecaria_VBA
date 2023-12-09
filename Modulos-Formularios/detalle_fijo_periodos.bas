Attribute VB_Name = "detalle_fijo_periodos"
Option Explicit

Sub detalle_fijo_periodos_varios()

Dim capital_inicial As Double
Dim plazos As Integer
Dim total_intereses As Double

Dim amortizacion_cuota_fija1 As Double
Dim cuota_mes_fija1 As Double
Dim capital_pendiente_fijo1 As Double
Dim periodo_n1 As Integer
Dim interes_n1 As Double
Dim j1_fijo As Double
Dim i1_fijo As Double
Dim z1 As Integer
Dim n_cuota1 As Integer
Dim intereses_cuota_fija1 As Double
Dim total_intereses1 As Double

Dim amortizacion_cuota_fija2 As Double
Dim cuota_mes_fija2 As Double
Dim capital_pendiente_fijo2 As Double
Dim periodo_n2 As Integer
Dim interes_n2 As Double
Dim j2_fijo As Double
Dim i2_fijo As Double
Dim z2 As Integer
Dim n_cuota2 As Integer
Dim intereses_cuota_fija2 As Double
Dim total_intereses2 As Double

plazos = Worksheets("formulario_fijo").Range("B1").Value
capital_pendiente_fijo1 = Worksheets("formulario_fijo").Range("B2").Value
periodo_n1 = Worksheets("formulario_fijo").Range("B4").Value
interes_n1 = Worksheets("formulario_fijo").Range("B5").Value
periodo_n2 = Worksheets("formulario_fijo").Range("B6").Value
interes_n2 = Worksheets("formulario_fijo").Range("B7").Value

total_intereses = 0


Rem aqui calculamos el primer periodo
               
    i1_fijo = interes_n1 / 1200
    j1_fijo = (1 + i1_fijo) ^ (-(plazos))
    cuota_mes_fija1 = (capital_pendiente_fijo1 * i1_fijo) / (1 - j1_fijo)
    
     For z1 = 1 To periodo_n1
               intereses_cuota_fija1 = (capital_pendiente_fijo1 * interes_n1) / 1200
               amortizacion_cuota_fija1 = cuota_mes_fija1 - intereses_cuota_fija1
               capital_pendiente_fijo1 = capital_pendiente_fijo1 - amortizacion_cuota_fija1
               n_cuota1 = z1
                                                                        
                                   Worksheets("cuadro_amortizacion_fijo").Select
                                   Worksheets("cuadro_amortizacion_fijo").Cells(1 + z1, 1).Value = n_cuota1
                                   Worksheets("cuadro_amortizacion_fijo").Cells(1 + z1, 2).Value = cuota_mes_fija1
                                   Worksheets("cuadro_amortizacion_fijo").Cells(1 + z1, 3).Value = intereses_cuota_fija1
                                   Worksheets("cuadro_amortizacion_fijo").Cells(1 + z1, 4).Value = amortizacion_cuota_fija1
                                   Worksheets("cuadro_amortizacion_fijo").Cells(1 + z1, 5).Value = capital_pendiente_fijo1
                                   
              total_intereses1 = total_intereses1 + intereses_cuota_fija1
                                   
     Next z1
     
capital_pendiente_fijo2 = capital_pendiente_fijo1
     
  Rem aqui calculamos el segundo periodo

               
    i2_fijo = interes_n2 / 1200
    j2_fijo = (1 + i2_fijo) ^ (-(plazos - periodo_n1))
    cuota_mes_fija2 = (capital_pendiente_fijo2 * i2_fijo) / (1 - j2_fijo)
    
     For z2 = (periodo_n1 + 1) To plazos
               intereses_cuota_fija2 = (capital_pendiente_fijo2 * interes_n2) / 1200
               amortizacion_cuota_fija2 = cuota_mes_fija2 - intereses_cuota_fija2
               capital_pendiente_fijo2 = capital_pendiente_fijo2 - amortizacion_cuota_fija2
               n_cuota2 = z2
                                                                        
                                   Worksheets("cuadro_amortizacion_fijo").Select
                                   Worksheets("cuadro_amortizacion_fijo").Cells(1 + z2, 1).Value = n_cuota2
                                   Worksheets("cuadro_amortizacion_fijo").Cells(1 + z2, 2).Value = cuota_mes_fija2
                                   Worksheets("cuadro_amortizacion_fijo").Cells(1 + z2, 3).Value = intereses_cuota_fija2
                                   Worksheets("cuadro_amortizacion_fijo").Cells(1 + z2, 4).Value = amortizacion_cuota_fija2
                                   Worksheets("cuadro_amortizacion_fijo").Cells(1 + z2, 5).Value = capital_pendiente_fijo2
                                   
              total_intereses2 = total_intereses2 + intereses_cuota_fija2
                                   
     Next z2
        
     total_intereses = total_intereses1 + total_intereses2
     
     Worksheets("formulario_fijo").Range("B8").Value = cuota_mes_fija1
     Worksheets("formulario_fijo").Range("B9").Value = cuota_mes_fija2
     Worksheets("formulario_fijo").Range("B10").Value = total_intereses
     Worksheets("formulario_fijo").Range("B11").Value = ((100 / Worksheets("formulario_fijo").Range("B2").Value)) * total_intereses
     
End Sub







