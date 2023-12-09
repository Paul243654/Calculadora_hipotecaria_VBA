Attribute VB_Name = "z7_listado_interes_legal_excel"
Sub exportar_SiNo()

Dim Pregunta As String

Pregunta = MsgBox("Desea expotar los resultados por cuota a un archivo excel", vbYesNo + vbQuestion, "EXPORTAR ARCHIVO")

    If Pregunta = vbYes Then
    
        MsgBox "El archivo se exportara en la misma carpeta donde esta guardado el ejecutable"
        Call cmd_exportarexcel_intlegal
    Else
    
        MsgBox "Elegiste no"

    End If

End Sub

Public Sub cmd_exportarexcel_intlegal()

Dim Fecha As String
Dim ruta As String
Dim titulo As String
Dim ultimo As Integer

On Error Resume Next

Worksheets("datos_con_int_legal").Select
ultimo = Worksheets("datos_con_int_legal").Range("A" & Rows.Count).End(xlUp).Row
Worksheets("datos_con_int_legal").Range(Cells(1, 1), Cells(ultimo, 18)).Copy

Workbooks.Add  'añadimos nuevo libro
ActiveSheet.Range("A1").PasteSpecial Paste:=xlPasteAll  'pegado espacial de excel de valores y formato, solo valores seria xl pastevalues

ruta = ThisWorkbook.Path
ruta = ruta & "\"
titulo = "Cuadro amortización mensual con interés legal"
Fecha = Now ' ahora remplazamos signos para qjue no den problemas a la hora de guardar
Fecha = Replace(Fecha, "/", "-") 'slash remplazado por guión
Fecha = Replace(Fecha, ":", ".")

Application.DisplayAlerts = False

Rem indicamos la ruta donde se guardara el libro nuevo
ActiveWorkbook.SaveAs Filename:=ruta & titulo & Fecha & ".xlsx"
ActiveWorkbook.Close

Application.DisplayAlerts = True

MsgBox "Archivo descargado", vbOKOnly, "Información"


End Sub
