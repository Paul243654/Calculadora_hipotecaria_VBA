Attribute VB_Name = "tabla_dinamica"
Option Explicit

Sub crear_tabla_dinamica()

Dim datos_tabla As Worksheet
Dim tdp As Worksheet
Dim PTcache As PivotCache
Dim Tabladinámica As PivotTable
Dim Rangodatos As Range
Dim últimafila As Long


Rem borra la tabla dinámica que se encuentra en la hoja dinámica

For Each Tabladinámica In Worksheets("tdp").PivotTables
        Tabladinámica.TableRange2.Clear
Next Tabladinámica

Rem definir el área de entrada y establecer el cache dinámico

últimafila = Worksheets("datos_tabla").Cells(Rows.Count, 1).End(xlUp).Row

Set Rangodatos = Worksheets("datos_tabla").Cells(1, 1).Resize(últimafila, 7)

Rem nos situamos en la hoja con los datos
'definimos la variable PTcache como valor intermedio necesario para la creación de la tabla dinámica

Sheets("datos_tabla").Select

Set PTcache = ActiveWorkbook.PivotCaches.Add(SourceType:=xlDatabase, SourceData:=Rangodatos.Address)

Rem se crea una tabla dinámica en blanco
'especificacmos la ubicación de salida y el nombre de la tabla

Set Tabladinámica = PTcache.CreatePivotTable(tabledestination:=Worksheets("tdp").Range("A1"), tablename:="pivottable1")

Rem se aplica el formato predefinido

Tabladinámica.Format xlReport6

Rem actualización automática

Tabladinámica.ManualUpdate = True

Tabladinámica.AddFields RowFields:=Array("nrev")


With Tabladinámica.PivotFields("ncuota")
.Orientation = xlDataField
.Function = xlCount
.Position = 1
.Caption = "Total cuotas"
End With

With Tabladinámica.PivotFields("irph")
.Orientation = xlDataField
.Function = xlAverage
.Position = 2
.Caption = "Valor irph"
End With

With Tabladinámica.PivotFields("euribor")
.Orientation = xlDataField
.Function = xlAverage
.Position = 3
.Caption = "Valor euribor"
End With

With Tabladinámica.PivotFields("nrev")
.Orientation = xlDataField
.Function = xlAverage
.Position = 4
.Caption = "nº de Revisión"
End With

With Tabladinámica.PivotFields("cuota_irph")
.Orientation = xlDataField
.Function = xlSum
.Position = 5
.Caption = "Acumulado cuota IRPH"
End With

With Tabladinámica.PivotFields("cuota_eur")
.Orientation = xlDataField
.Function = xlSum
.Position = 6
.Caption = "Acumulado cuota euribor"
End With

With Tabladinámica.PivotFields("dif_cuotas")
.Orientation = xlDataField
.Function = xlSum
.Position = 7
.Caption = "Diferencia cuotas"
End With

Tabladinámica.ManualUpdate = False

'Sheets("tdp").Select
'
'Sheets("tdp").CheckBox1.Value = True
'Sheets("tdp").CheckBox2.Value = True
'Sheets("tdp").CheckBox3.Value = True
'Sheets("tdp").CheckBox4.Value = True
'Sheets("tdp").CheckBox5.Value = True
'Sheets("tdp").CheckBox6.Value = True
'Sheets("tdp").CheckBox7.Value = True
'Sheets("tdp").CheckBox8.Value = True


End Sub

