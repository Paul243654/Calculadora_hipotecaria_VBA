Attribute VB_Name = "parte_1"
Option Explicit
Public capital_pendiente_actualizada As Variant

Sub años_plazo_fijo()

Dim año_hipoteca As Integer
Dim añost_hipoteca As Integer
Dim mes_revision As Integer
Dim primer_pago_mes As Integer
Dim capital_inicial As Double
Dim diferencial As Double
Dim año_actual As Integer
Dim mes_actual As Integer
Dim plazos As Integer
Dim diferencial_sustitutivo As Double

Dim amortizacion_cuota_irph As Double
Dim amortizacion_cuota_euribor As Double
Dim interes_irph_cuota As Double
Dim interes_euribor_cuota As Double
Dim cuota_mes_irph As Double
Dim cuota_mes_euribor As Double
Dim numero_cuotas As Integer
Dim capital_pendiente_irph As Double
Dim capital_pendiente_euribor As Double
Dim diferencia_entre_cuotas As Double
Dim cantidad_destinada_amortizar As Double
Dim diferencia_entre_capitales_pendientes As Double
Dim cantidad_a_devolver As Double

Dim dato_irph_año As Double
Dim dato_euribor_año As Double
Dim tabla_irph As Range
Dim tabla_euribor As Range

Dim buscarv_irph As Double
Dim buscarv_euribor As Double
Dim interes_irph_con_diferencial As Double
Dim interes_euribor_con_diferencial As Double

Dim z As Integer
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
Dim e As Integer
Dim f As Integer

Dim j_irph As Double
Dim j_euribor As Double
Dim i_irph As Double
Dim i_euribor As Double
Dim n As Integer

Dim uf As Integer
Dim ndemes As Integer

Dim años_plazo_fijo As Integer
Dim interes_fijo As Double
Dim revision As Integer
Dim n_revision As Integer

año_hipoteca = Worksheets("formulario").Range("B1").Value
añost_hipoteca = Worksheets("formulario").Range("B2").Value
mes_revision = Worksheets("formulario").Range("B3").Value
primer_pago_mes = Worksheets("formulario").Range("B4").Value
capital_inicial = Worksheets("formulario").Range("B5").Value
diferencial = Worksheets("formulario").Range("B6").Value
año_actual = Worksheets("formulario").Range("B7").Value
mes_actual = Worksheets("formulario").Range("B8").Value
diferencial_sustitutivo = Worksheets("formulario").Range("B11").Value
Set tabla_irph = Worksheets("datos_interes").Range("A2:M29")
Set tabla_euribor = Worksheets("datos_interes").Range("N2:Z29")
plazos = añost_hipoteca * 12
años_plazo_fijo = Worksheets("formulario").Range("B9").Value
revision = año_actual - año_hipoteca

 capital_pendiente_irph = capital_inicial
 capital_pendiente_euribor = capital_inicial
  
  For e = año_hipoteca To (año_hipoteca + años_plazo_fijo - 1)
  
   uf = Worksheets("cuadro_amortizacion").Range("A" & Rows.Count).End(xlUp).Row
    
    interes_fijo = Worksheets("formulario").Range("B10").Value
               
    i_irph = interes_fijo / 1200
    i_euribor = interes_fijo / 1200
               
    j_irph = (1 + i_irph) ^ (-(plazos))
    j_euribor = (1 + i_euribor) ^ (-(plazos))
               
     cuota_mes_irph = (capital_pendiente_irph * i_irph) / (1 - j_irph)
     cuota_mes_euribor = (capital_pendiente_euribor * i_euribor) / (1 - j_euribor)

     ndemes = primer_pago_mes
     n_revision = revision - (año_actual - e - 1)

     For z = 1 To 12

               
               interes_irph_cuota = (capital_pendiente_irph * interes_fijo) / 1200
               interes_euribor_cuota = (capital_pendiente_euribor * interes_fijo) / 1200
               
               amortizacion_cuota_irph = cuota_mes_irph - interes_irph_cuota
               amortizacion_cuota_euribor = cuota_mes_euribor - interes_euribor_cuota
               
               capital_pendiente_irph = capital_pendiente_irph - amortizacion_cuota_irph
               capital_pendiente_euribor = capital_pendiente_euribor - amortizacion_cuota_euribor
               
               diferencia_entre_cuotas = cuota_mes_irph - cuota_mes_euribor
               cantidad_destinada_amortizar = amortizacion_cuota_euribor - amortizacion_cuota_irph
               cantidad_a_devolver = (interes_irph_cuota - interes_euribor_cuota) - Abs(cantidad_destinada_amortizar)
               
               If cantidad_destinada_amortizar < 0 Then
               
                    cantidad_destinada_amortizar = 0
                    
               End If
               
                                                                        
                                   Worksheets("cuadro_amortizacion").Select
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 1).Value = n_revision
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 2).Value = año_hipoteca
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 3).Value = ndemes
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 4).Value = uf + z - 1
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 5).Value = cuota_mes_irph
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 6).Value = interes_irph_cuota
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 7).Value = amortizacion_cuota_irph
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 8).Value = capital_pendiente_irph
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 9).Value = interes_fijo
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 10).Value = cuota_mes_euribor
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 11).Value = interes_euribor_cuota
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 12).Value = amortizacion_cuota_euribor
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 13).Value = capital_pendiente_euribor
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 14).Value = interes_fijo
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 15).Value = diferencia_entre_cuotas
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 16).Value = cantidad_destinada_amortizar
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 17).Value = cantidad_a_devolver
                                   Worksheets("cuadro_amortizacion").Cells(uf + z, 18).Value = ndemes & año_hipoteca
                                   
                                   
                                   If ndemes = 12 Then
                                   ndemes = 1
                                   año_hipoteca = año_hipoteca + 1
                                   Else
                                   ndemes = ndemes + 1
                                   año_hipoteca = año_hipoteca
                                   End If
                                   
                                   
     Next z
     
     plazos = plazos - 12
     

Next e

capital_pendiente_actualizada = capital_pendiente_irph
capital_pendiente_actualizada = capital_pendiente_irph

End Sub

