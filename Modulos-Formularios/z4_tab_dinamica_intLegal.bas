Attribute VB_Name = "z4_tab_dinamica_intLegal"
Sub crear_tabla_dinamicaLegal()

Rem #6

Dim datos_volcados As Worksheet
Dim tabla_int_legal As Worksheet
Dim PTcache As PivotCache
Dim TabladinámicaLegal As PivotTable
Dim RangodatosLegal As Range
Dim últimafilaLegal As Long


Rem borra la tabla dinámicaLegal que se encuentra en la hoja dinámica

For Each TabladinámicaLegal In Worksheets("tabla_int_legal").PivotTables
        TabladinámicaLegal.TableRange2.Clear
Next TabladinámicaLegal

Rem definir el área de entrada y establecer el cache dinámico

últimafilaLegal = Worksheets("datos_volcados").Cells(Rows.Count, 1).End(xlUp).Row

Set RangodatosLegal = Worksheets("datos_volcados").Cells(1, 1).Resize(últimafilaLegal, 8)

Rem nos situamos en la hoja con los datos
'definimos la variable PTcache como valor intermedio necesario para la creación de la tabla dinámica

Sheets("datos_volcados").Select

Set PTcache = ActiveWorkbook.PivotCaches.Add(SourceType:=xlDatabase, SourceData:=RangodatosLegal.Address)

Rem se crea una tabla dinámica en blanco
'especificacmos la ubicación de salida y el nombre de la tabla

Set TabladinámicaLegal = PTcache.CreatePivotTable(tabledestination:=Worksheets("tabla_int_legal").Range("A1"), tablename:="pivottable1")

Rem se aplica el formato predefinido

TabladinámicaLegal.Format xlReport6

Rem actualización automática

TabladinámicaLegal.ManualUpdate = True

TabladinámicaLegal.AddFields RowFields:=Array("ncuota")



With TabladinámicaLegal.PivotFields("ndias")
.Orientation = xlDataField
.Function = xlSum
.Position = 1
.Caption = "n_días"
End With

With TabladinámicaLegal.PivotFields("InteresLegal")
.Orientation = xlDataField
.Function = xlSum
.Position = 2
.Caption = "Interes_Legal"
End With



TabladinámicaLegal.ManualUpdate = False


End Sub

