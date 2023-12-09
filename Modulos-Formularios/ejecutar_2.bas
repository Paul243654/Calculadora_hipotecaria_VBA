Attribute VB_Name = "ejecutar_2"
Option Explicit

Sub ejecutar_calculo_2()

Dim años_interes_fijo As Integer
Dim primer_pago_mes As Integer
Dim mes_revision As Integer

mes_revision = Worksheets("formulario").Range("B3").Value
primer_pago_mes = Worksheets("formulario").Range("B4").Value
años_interes_fijo = Worksheets("formulario").Range("B9").Value


If primer_pago_mes < mes_revision Then

          If años_interes_fijo > 0 Then
          
               Call parte_0_borrado_datos
               Call años_plazo_fijo
               Call mespago_menor_mesrevision_cancelada 'aqui crear otro modulo
               Call borrar_meses_demas
               Call format_cuadro_amortizacion
               Call resultados_finales
          
          Else
          
               Call parte_0_borrado_datos
               Call mespago_menor_mesrevision_cancelada 'aqui crear otro modulo
               Call borrar_meses_demas
               Call format_cuadro_amortizacion
               Call resultados_finales
          
          End If

Else

          If años_interes_fijo > 0 Then
          
               Call parte_0_borrado_datos
               Call años_plazo_fijo
               Call calculo_2_volcado_datos_cancelada 'aqui crear otro modulo
               Call borrar_meses_demas
               Call format_cuadro_amortizacion
               Call resultados_finales
          
          Else
          
               Call parte_0_borrado_datos
               Call calculo_2_volcado_datos_cancelada 'aqui crear otro modulo
               Call borrar_meses_demas
               Call format_cuadro_amortizacion
               Call resultados_finales
          
          End If
          
End If

Call copiar_datos_para_tabla
Call crear_tabla_dinamica
Call crear_grafico



End Sub
