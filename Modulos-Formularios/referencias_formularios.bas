Attribute VB_Name = "referencias_formularios"
Option Explicit

Sub declaración_nombres()

Worksheets("formulario").Select
Worksheets("formulario").Range("A1").Value = "Año de firma de la hipoteca"
Worksheets("formulario").Range("A2").Value = "Plazo (años)"
Worksheets("formulario").Range("A3").Value = "nº de mes de revisión"
Worksheets("formulario").Range("A4").Value = "nº de mes del 1er pago"
Worksheets("formulario").Range("A5").Value = "Capital inicial"
Worksheets("formulario").Range("A6").Value = "Diferencial"
Worksheets("formulario").Range("A7").Value = "Año en curso"
Worksheets("formulario").Range("A8").Value = "nº de mes en curso"
Worksheets("formulario").Range("A9").Value = "Años a plazo fijo"
Worksheets("formulario").Range("A10").Value = "Interés a plazo fijo"
Worksheets("formulario").Range("A11").Value = "Diferencial sustitutivo"

Worksheets("formulario_simulacion").Select
Worksheets("formulario_simulacion").Range("A1").Value = "Cuotas pendientes de pago"
Worksheets("formulario_simulacion").Range("A2").Value = "nº de mes de revisión"
Worksheets("formulario_simulacion").Range("A3").Value = "nº de mes del 1er pago"
Worksheets("formulario_simulacion").Range("A4").Value = "Capital pendiente"
Worksheets("formulario_simulacion").Range("A5").Value = "Diferencial sustitutivo"
Worksheets("formulario_simulacion").Range("A6").Value = "Último año de revisión"
Worksheets("formulario_simulacion").Range("A7").Value = "nº de mes en curso"
Worksheets("formulario_simulacion").Range("A9").Value = "Nueva cuota simulada "
Worksheets("formulario_simulacion").Range("A10").Value = "Amortización de cuota simulada"
Worksheets("formulario_simulacion").Range("A11").Value = "Interés de cuota simulada"


End Sub
