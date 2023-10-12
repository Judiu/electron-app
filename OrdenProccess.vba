Sub ProcesarArchivosExcel()
    Dim Carpeta As String
    Dim Archivo As String
    Dim Libro As Workbook
    
    ' Especifica la carpeta donde se encuentran los archivos de Excel
    Carpeta = "F:\LPS Ingenieria Estructural\LPS Ingenieria Estructural - Ferroscan\Informes Entrega Final\CR Novaterra Cra. 6A #14-37 Sur\Datos" ' Cambia la ruta a la carpeta adecuada
    
    ' Inicia un bucle para recorrer todos los archivos en la carpeta
    Archivo = Dir(Carpeta & "\*.csv") ' Cambia la extensión según tus archivos
    
    Do While Archivo <> ""
        ' Abre el archivo de Excel
        Set Libro = Workbooks.Open(Carpeta & "\" & Archivo)
        
        ' Ejecuta el macro que deseas aplicar al archivo
        ' Reemplaza "NombreDelMacro" con el nombre de tu macro
        Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=True, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
        ), Array(14, 1)), TrailingMinusNumbers:=True
    Range("B3:H3").Select
    Selection.Cut Destination:=Range("C3:I3")
    Range("G3:I3").Select
    Selection.Cut Destination:=Range("I3:K3")
    Range("I3:K3").Select
        
        Call Promedio2
        
        ' Cierra y guarda el archivo de Excel (si es necesario)
        Libro.Close SaveChanges:=True
        
        ' Continúa con el siguiente archivo en la carpeta
        Archivo = Dir
    Loop
End Sub
