Attribute VB_Name = "parte_3"
Option Explicit

Sub borrar_meses_demas()

Dim año_actual As Integer
Dim mes_actual As Integer
Dim dato_coincidir As Double
Dim posicion As Integer
Dim fila_a_eliminar As Integer

año_actual = Worksheets("formulario").Range("B7").Value
mes_actual = Worksheets("formulario").Range("B8").Value

dato_coincidir = mes_actual & año_actual

posicion = Application.Match(dato_coincidir, Worksheets("cuadro_amortizacion").Range("R2:R500"), 0)

fila_a_eliminar = posicion + 2

Worksheets("cuadro_amortizacion").Select
Worksheets("cuadro_amortizacion").Range(Cells(fila_a_eliminar, 1), Cells(fila_a_eliminar + 20, 18)).ClearContents


End Sub

