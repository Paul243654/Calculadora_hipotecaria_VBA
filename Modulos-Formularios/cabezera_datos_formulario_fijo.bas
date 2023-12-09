Attribute VB_Name = "cabezera_datos_formulario_fijo"
Option Explicit

Sub cabezera_formulario_fijo()

Worksheets("formulario_fijo").Select
Worksheets("formulario_fijo").Cells.ClearContents

Worksheets("formulario_fijo").Select
Worksheets("formulario_fijo").Range("A1").Value = "nº de Plazos"
Worksheets("formulario_fijo").Range("A2").Value = "Capital inicial (€)"
Worksheets("formulario_fijo").Range("A3").Value = "Interés a plazo fijo (%)"
Worksheets("formulario_fijo").Range("A4").Value = "1er periodo de plazos"
Worksheets("formulario_fijo").Range("A5").Value = "Interés del 1er periodo (%)"
Worksheets("formulario_fijo").Range("A6").Value = "2do periodo de plazos"
Worksheets("formulario_fijo").Range("A7").Value = "Interés del 2do periodo (%)"
Worksheets("formulario_fijo").Range("A8").Value = "Cuota 1er periodo (€)"
Worksheets("formulario_fijo").Range("A9").Value = "Cuota 2do periodo (€)"
Worksheets("formulario_fijo").Range("A10").Value = "Total intereses (€)"
Worksheets("formulario_fijo").Range("A11").Value = "interés pagado con respecto al total (%)"

End Sub

