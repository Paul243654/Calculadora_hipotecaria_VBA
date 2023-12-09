Attribute VB_Name = "nueva_cuota"
Option Explicit

Sub calculo_cuota_nueva()


Dim tabla_euribor As Range
Dim buscarv_euribor As Double
Dim interes_euribor_con_diferencial As Double
Dim i_euribor As Double
Dim j_euribor As Double
Dim cuota_mes_euribor As Double
Dim interes_euribor_cuota As Double
Dim amortizacion_cuota_euribor As Double
Dim ultimo_año_revision As Integer
Dim mes_revision As Integer
Dim plazos As Integer
Dim diferencial_sustitutivo_simulacion As Double
Dim capital_pendiente_euribor As Double
Dim uf As Double
Dim ufg As Double

uf = Worksheets("cuadro_amortizacion").Cells(Rows.Count, 1).End(xlUp).Row
ufg = uf + 1

Worksheets("formulario_simulacion").Range("B7").Value = Worksheets("formulario").Range("B8").Value
Worksheets("formulario_simulacion").Range("B4").Value = (Worksheets("formulario").Range("B5").Value) - (Worksheets("cuadro_amortizacion").Cells(ufg, 7).Value) - (Worksheets("cuadro_amortizacion").Cells(ufg, 16).Value)
Worksheets("formulario_simulacion").Range("B3").Value = Worksheets("formulario").Range("B4").Value
Worksheets("formulario_simulacion").Range("B2").Value = Worksheets("formulario").Range("B3").Value
Worksheets("formulario_simulacion").Range("B1").Value = ((Worksheets("formulario").Range("B2").Value) * 12) - (Worksheets("cuadro_amortizacion").Cells(ufg, 4).Value)


ultimo_año_revision = Worksheets("formulario_simulacion").Range("B6").Value
mes_revision = Worksheets("formulario_simulacion").Range("B2").Value
plazos = Worksheets("formulario_simulacion").Range("B1").Value
capital_pendiente_euribor = Worksheets("formulario_simulacion").Range("B4").Value
diferencial_sustitutivo_simulacion = Worksheets("formulario_simulacion").Range("B5").Value

Set tabla_euribor = Worksheets("datos_interes").Range("N2:Z29")

buscarv_euribor = Application.VLookup(ultimo_año_revision, tabla_euribor, mes_revision + 1, False)
interes_euribor_con_diferencial = buscarv_euribor + diferencial_sustitutivo_simulacion

    i_euribor = interes_euribor_con_diferencial / 1200

    j_euribor = (1 + i_euribor) ^ (-(plazos))

    cuota_mes_euribor = (capital_pendiente_euribor * i_euribor) / (1 - j_euribor)

    interes_euribor_cuota = (capital_pendiente_euribor * interes_euribor_con_diferencial) / 1200

    amortizacion_cuota_euribor = cuota_mes_euribor - interes_euribor_cuota
                                   
                         
Worksheets("formulario_simulacion").Range("B9").Value = cuota_mes_euribor
Worksheets("formulario_simulacion").Range("B9").NumberFormat = "#,##0.00"
Worksheets("formulario_simulacion").Range("B10").Value = amortizacion_cuota_euribor
Worksheets("formulario_simulacion").Range("B10").NumberFormat = "#,##0.00"
Worksheets("formulario_simulacion").Range("B11").Value = interes_euribor_cuota
Worksheets("formulario_simulacion").Range("B11").NumberFormat = "#,##0.00"
                                                  
                         
                         
                         
End Sub


