Attribute VB_Name = "datos_evolucion_indice"
Option Explicit

Sub evolucion_indice_hipotecario()

Dim año_hipoteca As Integer
Dim año_actual As Integer
Dim mes_revision As Integer
Dim ih As Integer
Dim tabla_irph As Range
Dim tabla_euribor As Range
Dim i As Integer
Dim buscarv_irph As Double
Dim buscarv_euribor As Double
Dim diferencial_sustitutivo As Double


Worksheets("graf_evol").Cells.ClearContents
Worksheets("graf_evol").Select
ActiveSheet.ChartObjects.Delete
Worksheets("graf_evol").Range("A1").Value = " Año "
Worksheets("graf_evol").Range("B1").Value = " IRPH "
Worksheets("graf_evol").Range("C1").Value = " Euribor "

año_hipoteca = Worksheets("formulario").Range("B1").Value
año_actual = Worksheets("formulario").Range("B7").Value
mes_revision = Worksheets("formulario").Range("B3").Value
diferencial_sustitutivo = Worksheets("formulario").Range("B11").Value

Set tabla_irph = Worksheets("datos_interes").Range("A2:M29")
Set tabla_euribor = Worksheets("datos_interes").Range("N2:Z29")

i = 0

For ih = año_hipoteca To año_actual

 buscarv_irph = Application.VLookup(ih, tabla_irph, mes_revision + 1, False)
 
 If buscarv_irph = 0 Then
 Worksheets("graf_evol").Cells(2 + i, 1).Value = mes_revision & "_" & ih
 Worksheets("graf_evol").Cells(2 + i, 2).Value = Empty
 Else
 Worksheets("graf_evol").Cells(2 + i, 1).Value = mes_revision & "_" & ih
 Worksheets("graf_evol").Cells(2 + i, 2).Value = buscarv_irph
 End If
 
 buscarv_euribor = Application.VLookup(ih, tabla_euribor, mes_revision + 1, False)
 
 If buscarv_euribor = 0 Then
 Worksheets("graf_evol").Cells(2 + i, 3).Value = Empty
 Else
 Worksheets("graf_evol").Cells(2 + i, 3).Value = buscarv_euribor
 End If
 
 i = 1 + i

Next ih

End Sub

Sub formatear_graf_evol()

Dim ufm As Integer
Dim r As Integer
Dim s As Integer

Worksheets("graf_evol").Select

ufm = Worksheets("graf_evol").Range("A" & Rows.Count).End(xlUp).Row

For s = 2 To ufm
For r = 2 To 3

Cells(s, r).Select
ActiveCell.NumberFormat = "##0.000"

Next r
Next s

End Sub

Sub creacion_grafico_evolucion()

Dim ufi As Integer
Dim evolucion As ChartObject

ufi = Worksheets("graf_evol").Range("A" & Rows.Count).End(xlUp).Row

Set evolucion = Sheets("graf_evol").ChartObjects.Add(Left:=200, Top:=10, Width:=350, Height:=200)

With evolucion.Chart
.SetSourceData Source:=Sheets("graf_evol").Range(Cells(1, 1), Cells(ufi, 3))
.ChartType = xlLine
.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 170, 171)
.HasTitle = True
.ChartTitle.Text = "Evolución de indices"
End With

With evolucion.Chart.Axes(xlCategory, xlPrimary)
.HasTitle = True
.AxisTitle.Characters.Text = "Mes_Año"
Rem .HasMajorGridlines = True 'Líneas de grilla verticales
.TickLabels.Orientation = xlUpward
End With

With evolucion.Chart.Axes(xlValue, xlPrimary)
.HasTitle = True
.AxisTitle.Characters.Text = "valor índice (%)"


End With

With evolucion.Chart.PlotArea
.Width = evolucion.Chart.PlotArea.Width + 70
.Height = evolucion.Chart.PlotArea.Height + 80
.Left = 10
.Top = 10
End With


Rem .TickLabels.Orientation = xlUpward

End Sub
