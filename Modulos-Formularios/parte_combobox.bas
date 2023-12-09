Attribute VB_Name = "parte_combobox"
Option Explicit

Sub ejecutar_listado_combobox()

Dim i As Integer
Dim ñ As Integer
Dim s As Integer
Dim bc As Integer
Dim dk As Integer

Worksheets("combobox").Select


For i = 1 To 12

Worksheets("combobox").Cells(i, 1) = i

Next i

For ñ = 1 To 28

Worksheets("combobox").Cells(ñ, 2) = 1996 + ñ

Next ñ


For s = 1 To 40

Worksheets("combobox").Cells(s, 3) = s

Next s

For bc = 1 To 11
Worksheets("combobox").Cells(bc, 4) = bc
Next bc

For dk = 12 To 14
Worksheets("combobox").Cells(dk, 4) = 3 - dk
Next dk


End Sub
