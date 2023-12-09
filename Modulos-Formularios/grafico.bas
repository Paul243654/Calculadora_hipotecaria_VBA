Attribute VB_Name = "grafico"
Option Explicit

Sub crear_grafico()

Dim diferencias As ChartObject
Dim ufm As Long
Dim ufmn As Long
Dim td As Sheets
Dim datos_grafico As Sheets
Dim i As Integer
Dim j As Integer
Dim s As Integer
Dim r As Integer

On Error Resume Next
Worksheets("datos_grafico").Select
Worksheets("datos_grafico").Cells.ClearContents
ActiveSheet.ChartObjects.Delete

Worksheets("tdp").Select
ufm = Worksheets("tdp").Range("E" & Rows.Count).End(xlUp).Row
ufmn = ufm - 1

Worksheets("tdp").Select
Worksheets("tdp").Range(Cells(1, 2), Cells(ufmn, 2)).Select
Selection.Copy
Worksheets("datos_grafico").Select
Worksheets("datos_grafico").Range(Cells(1, 1), Cells(ufmn, 1)).PasteSpecial xlPasteAll

Worksheets("tdp").Select
Worksheets("tdp").Range(Cells(1, 3), Cells(ufmn, 3)).Select
Selection.Copy
Worksheets("datos_grafico").Select
Worksheets("datos_grafico").Range(Cells(1, 2), Cells(ufmn, 2)).PasteSpecial xlPasteAll

Worksheets("tdp").Select
Worksheets("tdp").Range(Cells(1, 4), Cells(ufmn, 4)).Select
Selection.Copy
Worksheets("datos_grafico").Select
Worksheets("datos_grafico").Range(Cells(1, 3), Cells(ufmn, 3)).PasteSpecial xlPasteAll

Worksheets("tdp").Select
Worksheets("tdp").Range(Cells(1, 5), Cells(ufmn, 5)).Select
Selection.Copy
Worksheets("datos_grafico").Select
Worksheets("datos_grafico").Range(Cells(1, 4), Cells(ufmn, 4)).PasteSpecial xlPasteAll

Worksheets("tdp").Select
Worksheets("tdp").Range(Cells(1, 6), Cells(ufmn, 6)).Select
Selection.Copy
Worksheets("datos_grafico").Select
Worksheets("datos_grafico").Range(Cells(1, 5), Cells(ufmn, 5)).PasteSpecial xlPasteAll

Worksheets("tdp").Select
Worksheets("tdp").Range(Cells(1, 7), Cells(ufmn, 7)).Select
Selection.Copy
Worksheets("datos_grafico").Select
Worksheets("datos_grafico").Range(Cells(1, 6), Cells(ufmn, 6)).PasteSpecial xlPasteAll

Worksheets("tdp").Select
Worksheets("tdp").Range(Cells(1, 8), Cells(ufmn, 8)).Select
Selection.Copy
Worksheets("datos_grafico").Select
Worksheets("datos_grafico").Range(Cells(1, 7), Cells(ufmn, 7)).PasteSpecial xlPasteAll

Rem voy a dar formato de 2 y 3 decimales para la presentación

Worksheets("datos_grafico").Select

For i = 2 To ufmn
For j = 2 To 3

Cells(i, j).Select
ActiveCell.NumberFormat = "#,##0.000"

Next j
Next i


For s = 2 To ufmn
For r = 5 To 7

Cells(s, r).Select
ActiveCell.NumberFormat = "#,##00.00"

Next r
Next s

Rem fin del formateo de decimales a las celdas

Set diferencias = Sheets("datos_grafico").ChartObjects.Add(Left:=660, Top:=30, Width:=300, Height:=200)

With diferencias.Chart
.SetSourceData Source:=Sheets("datos_grafico").Range(Cells(1, 4), Cells(ufmn, 6))
.ChartType = xlColumnClustered
.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 170, 171)
.HasTitle = True
.ChartTitle.Text = "Diferencia anual de cuotas (€)"
End With

With diferencias.Chart.Axes(xlCategory, xlPrimary)
.HasTitle = True
.AxisTitle.Characters.Text = "Revisión nº"
End With

With diferencias.Chart.Axes(xlValue, xlPrimary)
.HasTitle = True
.AxisTitle.Characters.Text = "Total anual (€)"
End With

With diferencias.Chart.PlotArea
.Width = diferencias.Chart.PlotArea.Width + 25
.Height = diferencias.Chart.PlotArea.Height + 27
.Left = 20
.Top = 25
End With


End Sub



