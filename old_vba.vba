Sub TransferirDatosChloride()
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim ultimaFilaOrigen As Long
    Dim i As Long
    Dim consecutivo As Long
    Dim conn As Object
    Dim rs As Object
    Dim sql As String
    Dim lastLabSampleID As String
    Dim parentID As String
    Dim foundCell As Range
    Dim targetRow As Long
    Dim counterAnalitos As Long
    Dim colorFondo As Long
    Dim extractedData As String

    ' Conexión a la base de datos
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Driver={SQL Server};Server=np:\\.\pipe\LOCALDB#F81255E4\tsql\query;Database=SRLSQL;Trusted_Connection=yes;"

    ' Consulta SQL para obtener el último LabSampleID
    sql = "SELECT TOP 1 [LabSampleID] FROM [SRLSQL].[dbo].[Sample_Tests] ORDER BY [LabSampleID] DESC"

    ' Ejecutar la consulta
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, conn

    ' Obtener el valor del último LabSampleID
    If Not rs.EOF Then
        lastLabSampleID = Left(rs.Fields(0).Value, Len(rs.Fields(0).Value) - 4) + 1
    Else
        MsgBox "No se encontraron registros en la tabla Sample_Tests.", vbExclamation
        Exit Sub
    End If

    ' Cerrar la conexión
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    ' Configurar las hojas de cálculo
    Set wsOrigen = ThisWorkbook.Sheets("Chain of Custody")
    Set wsDestino = ThisWorkbook.Sheets("Reporte")

    ' Colocar el lastLabSampleID en las celdas AI11 y AQ11, combinando las celdas
    wsDestino.Range("AI11:AQ11").Merge
    wsDestino.Cells(11, "AI").Value = lastLabSampleID

    ' Encontrar la última fila con datos en la hoja de origen
    ultimaFilaOrigen = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row

    ' Inicializar la fila de destino en la fila 14
    Dim filaDestino As Long
    filaDestino = 14
    
    targetRow = 0

    ' Inicializar el consecutivo
    consecutivo = 1

    ' Desactivar alertas para evitar mensajes de confirmación
    Application.DisplayAlerts = False
    
    ' Definir el color de fondo
    colorFondo = RGB(218, 238, 243)

    ' Iterar sobre cada fila de la hoja de origen desde la fila 7 hasta la 26
    For i = 7 To 26
        ' Leer los valores de la hoja de origen
        Dim fecha As Variant
        Dim muestra As Variant
        Dim sampleMatrix As Variant
        Dim hora As Variant
        
        ' Asegúrate de que el valor se interprete como una fecha
        If IsDate(wsOrigen.Cells(i, 3).Value) Then
            fecha = Format(CDate(wsOrigen.Cells(i, 3).Value), "dd/mm/yyyy")
        Else
            fecha = wsOrigen.Cells(i, 3).Value ' Mantener el valor original si no es una fecha válida
        End If
        hora = Format(wsOrigen.Cells(i, 4).Value, "hh:mm")
        muestra = wsOrigen.Cells(i, 5).Value
        sampleMatrix = wsOrigen.Cells(i, 6).Value
        
        ' Comprobar si fechaHora no está vacía, es una fecha válida, y la muestra no está vacía
        If Not IsEmpty(muestra) Then
            ' Inicializar la fila de destino Analytical Results
            If targetRow = 0 Then
                Set filaAnalyticalResults = wsDestino.Cells.Find(What:="****************************** Analytical Results ******************************", LookIn:=xlValues, LookAt:=xlWhole)
                targetRow = filaAnalyticalResults.Row + 2
            End If
            
            ' Insertar cuatro nuevas filas en la hoja de destino
            wsDestino.Rows(targetRow).Insert Shift:=xlDown
            wsDestino.Rows(targetRow + 1).Insert Shift:=xlDown
            wsDestino.Rows(targetRow + 2).Insert Shift:=xlDown
            wsDestino.Rows(targetRow + 3).Insert Shift:=xlDown

            ' Formatear la primera fila debajo de la cadena encontrada
            With wsDestino
                .Range("B" & targetRow).Value = "Cliente Sample ID: "
                .Range("B" & targetRow & ":K" & targetRow).Interior.Color = colorFondo
                .Range("B" & targetRow & ":K" & targetRow).Merge
                .Range("B" & targetRow & ":K" & targetRow).HorizontalAlignment = xlRight
                .Range("B" & targetRow & ":K" & targetRow).Font.Bold = False
                .Range("B" & targetRow & ":K" & targetRow).Font.Size = 10
                
                .Range("L" & targetRow).Value = muestra
                .Range("L" & targetRow & ":S" & targetRow).Interior.Color = colorFondo
                .Range("L" & targetRow & ":S" & targetRow).Merge
                .Range("L" & targetRow & ":S" & targetRow).HorizontalAlignment = xlLeft
                .Range("L" & targetRow & ":S" & targetRow).Font.Bold = True
                .Range("L" & targetRow & ":S" & targetRow).Font.Size = 10
                
                .Range("T" & targetRow).Value = "Date Collected: "
                .Range("T" & targetRow & ":Y" & targetRow).Interior.Color = colorFondo
                .Range("T" & targetRow & ":Y" & targetRow).Merge
                .Range("T" & targetRow & ":Y" & targetRow).HorizontalAlignment = xlRight
                .Range("T" & targetRow & ":Y" & targetRow).Font.Bold = False
                .Range("T" & targetRow & ":Y" & targetRow).Font.Size = 10
                
                .Range("Z" & targetRow).Value = fecha & " " & hora
                .Range("Z" & targetRow & ":AC" & targetRow).Interior.Color = colorFondo
                .Range("Z" & targetRow & ":AC" & targetRow).Merge
                .Range("Z" & targetRow & ":AC" & targetRow).HorizontalAlignment = xlLeft
                .Range("Z" & targetRow & ":AC" & targetRow).Font.Bold = True
                .Range("Z" & targetRow & ":AC" & targetRow).Font.Size = 10
                
                .Range("AD" & targetRow).Value = "MATRIX ID: "
                .Range("AD" & targetRow & ":AH" & targetRow).Interior.Color = colorFondo
                .Range("AD" & targetRow & ":AH" & targetRow).Merge
                .Range("AD" & targetRow & ":AH" & targetRow).HorizontalAlignment = xlRight
                .Range("AD" & targetRow & ":AH" & targetRow).Font.Bold = False
                .Range("AD" & targetRow & ":AH" & targetRow).Font.Size = 10
                
                .Range("AI" & targetRow).Value = sampleMatrix
                .Range("AI" & targetRow & ":AQ" & targetRow).Interior.Color = colorFondo
                .Range("AI" & targetRow & ":AQ" & targetRow).Merge
                .Range("AI" & targetRow & ":AQ" & targetRow).HorizontalAlignment = xlLeft
                .Range("AI" & targetRow & ":AQ" & targetRow).Font.Bold = True
                .Range("AI" & targetRow & ":AQ" & targetRow).Font.Size = 10
            End With

            targetRow2 = targetRow + 1

            ' Formatear la segunda fila
            With wsDestino
                .Range("B" & targetRow2).Value = "Lab Sample ID: "
                .Range("B" & targetRow2 & ":K" & targetRow2).Interior.Color = colorFondo
                .Range("B" & targetRow2 & ":K" & targetRow2).Font.Bold = False
                .Range("B" & targetRow2 & ":K" & targetRow2).Merge
                .Range("B" & targetRow2 & ":K" & targetRow2).HorizontalAlignment = xlRight
                .Range("B" & targetRow2 & ":K" & targetRow2).Font.Size = 10
                
                .Range("L" & targetRow2).Value = parentID
                .Range("L" & targetRow2 & ":S" & targetRow2).Interior.Color = colorFondo
                .Range("L" & targetRow2 & ":S" & targetRow2).Merge
                .Range("L" & targetRow2 & ":S" & targetRow2).HorizontalAlignment = xlLeft
                .Range("L" & targetRow2 & ":S" & targetRow2).Font.Bold = True
                .Range("L" & targetRow2 & ":S" & targetRow2).Font.Size = 10
                
                .Range("T" & targetRow2).Value = "Collected By: "
                .Range("T" & targetRow2 & ":Y" & targetRow2).Interior.Color = colorFondo
                .Range("T" & targetRow2 & ":Y" & targetRow2).Merge
                .Range("T" & targetRow2 & ":Y" & targetRow2).Font.Bold = False
                .Range("T" & targetRow2 & ":Y" & targetRow2).HorizontalAlignment = xlRight
                .Range("T" & targetRow2 & ":Y" & targetRow2).Font.Size = 10
                
                .Range("Z" & targetRow2).Value = wsOrigen.Range("G4").Value
                .Range("Z" & targetRow2 & ":AC" & targetRow2).Interior.Color = colorFondo
                .Range("Z" & targetRow2 & ":AC" & targetRow2).Merge
                .Range("Z" & targetRow2 & ":AC" & targetRow2).HorizontalAlignment = xlLeft
                .Range("Z" & targetRow2 & ":AC" & targetRow2).Font.Bold = True
                .Range("Z" & targetRow2 & ":AC" & targetRow2).Font.Size = 10
                
                .Range("AD" & targetRow2).Value = ""
                .Range("AD" & targetRow2 & ":AH" & targetRow2).Interior.Color = colorFondo
                .Range("AD" & targetRow2 & ":AH" & targetRow2).Merge
                .Range("AD" & targetRow2 & ":AH" & targetRow2).HorizontalAlignment = xlRight
                .Range("AD" & targetRow2 & ":AH" & targetRow2).Font.Size = 10
                
                .Range("AI" & targetRow2).Value = ""
                .Range("AI" & targetRow2 & ":AQ" & targetRow2).Interior.Color = colorFondo
                .Range("AI" & targetRow2 & ":AQ" & targetRow2).Merge
                .Range("AI" & targetRow2 & ":AQ" & targetRow2).HorizontalAlignment = xlLeft
                .Range("AI" & targetRow2 & ":AQ" & targetRow2).Font.Size = 10
            End With
            
            targetRow3 = targetRow2 + 1

            ' Formatear la tercera fila
            With wsDestino
                .Range("C" & targetRow3).Value = "Classical Chemistry Parameters"
                .Range("B" & targetRow3 & ":AQ" & targetRow3).Interior.Color = colorFondo
                .Range("C" & targetRow3 & ":AQ" & targetRow3).Merge
                .Range("C" & targetRow3 & ":AQ" & targetRow3).HorizontalAlignment = xlLeft
                .Range("C" & targetRow3 & ":AQ" & targetRow3).Font.Bold = True
                .Range("C" & targetRow3 & ":AQ" & targetRow3).Font.Size = 10
            End With
            
            targetRow4 = targetRow3 + 1

            ' Formatear la cuarta fila
            With wsDestino
                .Range("B" & targetRow4).Value = "Analyte Name (Analyte ID)"
                .Range("B" & targetRow4 & ":J" & targetRow4).Interior.Color = RGB(255, 255, 255) ' Fondo blanco
                .Range("B" & targetRow4 & ":J" & targetRow4).Merge
                .Range("B" & targetRow4 & ":J" & targetRow4).HorizontalAlignment = xlCenter
                .Range("B" & targetRow4 & ":J" & targetRow4).Font.Bold = True
                .Range("B" & targetRow4 & ":J" & targetRow4).Font.Size = 10
                .Range("B" & targetRow4 & ":J" & targetRow4).Borders(xlEdgeTop).LineStyle = xlContinuous ' Borde superior
                .Range("B" & targetRow4 & ":J" & targetRow4).Borders(xlEdgeBottom).LineStyle = xlContinuous ' Borde inferior
                .Range("B" & targetRow4 & ":J" & targetRow4).Borders(xlEdgeBottom).Weight = xlMedium ' Borde inferior grueso
            
                .Range("K" & targetRow4).Value = "Results/Qual"
                .Range("K" & targetRow4 & ":S" & targetRow4).Interior.Color = RGB(255, 255, 255) ' Fondo blanco
                .Range("K" & targetRow4 & ":S" & targetRow4).Merge
                .Range("K" & targetRow4 & ":S" & targetRow4).HorizontalAlignment = xlCenter
                .Range("K" & targetRow4 & ":S" & targetRow4).Font.Bold = True
                .Range("K" & targetRow4 & ":S" & targetRow4).Font.Size = 10
                .Range("K" & targetRow4 & ":S" & targetRow4).Borders(xlEdgeTop).LineStyle = xlContinuous ' Borde superior
                .Range("K" & targetRow4 & ":S" & targetRow4).Borders(xlEdgeBottom).LineStyle = xlContinuous ' Borde inferior
                .Range("K" & targetRow4 & ":S" & targetRow4).Borders(xlEdgeBottom).Weight = xlMedium ' Borde inferior grueso
            
                .Range("T" & targetRow4).Value = "Units"
                .Range("T" & targetRow4 & ":U" & targetRow4).Interior.Color = RGB(255, 255, 255) ' Fondo blanco
                .Range("T" & targetRow4 & ":U" & targetRow4).Merge
                .Range("T" & targetRow4 & ":U" & targetRow4).HorizontalAlignment = xlCenter
                .Range("T" & targetRow4 & ":U" & targetRow4).Font.Bold = True
                .Range("T" & targetRow4 & ":U" & targetRow4).Font.Size = 10
                .Range("T" & targetRow4 & ":U" & targetRow4).Borders(xlEdgeTop).LineStyle = xlContinuous ' Borde superior
                .Range("T" & targetRow4 & ":U" & targetRow4).Borders(xlEdgeBottom).LineStyle = xlContinuous ' Borde inferior
                .Range("T" & targetRow4 & ":U" & targetRow4).Borders(xlEdgeBottom).Weight = xlMedium ' Borde inferior grueso
            
                .Range("V" & targetRow4).Value = "DF"
                .Range("V" & targetRow4 & ":X" & targetRow4).Interior.Color = RGB(255, 255, 255) ' Fondo blanco
                .Range("V" & targetRow4 & ":X" & targetRow4).Merge
                .Range("V" & targetRow4 & ":X" & targetRow4).HorizontalAlignment = xlCenter
                .Range("V" & targetRow4 & ":X" & targetRow4).Font.Bold = True
                .Range("V" & targetRow4 & ":X" & targetRow4).Font.Size = 10
                .Range("V" & targetRow4 & ":X" & targetRow4).Borders(xlEdgeTop).LineStyle = xlContinuous ' Borde superior
                .Range("V" & targetRow4 & ":X" & targetRow4).Borders(xlEdgeBottom).LineStyle = xlContinuous ' Borde inferior
                .Range("V" & targetRow4 & ":X" & targetRow4).Borders(xlEdgeBottom).Weight = xlMedium ' Borde inferior grueso
            
                .Range("Y" & targetRow4).Value = "MDL"
                .Range("Y" & targetRow4 & ":AA" & targetRow4).Interior.Color = RGB(255, 255, 255) ' Fondo blanco
                .Range("Y" & targetRow4 & ":AA" & targetRow4).Merge
                .Range("Y" & targetRow4 & ":AA" & targetRow4).HorizontalAlignment = xlCenter
                .Range("Y" & targetRow4 & ":AA" & targetRow4).Font.Bold = True
                .Range("Y" & targetRow4 & ":AA" & targetRow4).Font.Size = 10
                .Range("Y" & targetRow4 & ":AA" & targetRow4).Borders(xlEdgeTop).LineStyle = xlContinuous ' Borde superior
                .Range("Y" & targetRow4 & ":AA" & targetRow4).Borders(xlEdgeBottom).LineStyle = xlContinuous ' Borde inferior
                .Range("Y" & targetRow4 & ":AA" & targetRow4).Borders(xlEdgeBottom).Weight = xlMedium ' Borde inferior grueso
            
                .Range("AB" & targetRow4).Value = "PQL"
                .Range("AB" & targetRow4).Interior.Color = RGB(255, 255, 255) ' Fondo blanco
                .Range("AB" & targetRow4).Merge
                .Range("AB" & targetRow4).HorizontalAlignment = xlCenter
                .Range("AB" & targetRow4).Font.Bold = True
                .Range("AB" & targetRow4).Font.Size = 10
                .Range("AB" & targetRow4).Borders(xlEdgeTop).LineStyle = xlContinuous ' Borde superior
                .Range("AB" & targetRow4).Borders(xlEdgeBottom).LineStyle = xlContinuous ' Borde inferior
                .Range("AB" & targetRow4).Borders(xlEdgeBottom).Weight = xlMedium ' Borde inferior grueso
            
                .Range("AC" & targetRow4).Value = "Method Analyzed"
                .Range("AC" & targetRow4).Interior.Color = RGB(255, 255, 255) ' Fondo blanco
                .Range("AC" & targetRow4).HorizontalAlignment = xlCenter
                .Range("AC" & targetRow4).Font.Bold = True
                .Range("AC" & targetRow4).Font.Size = 10
                .Range("AC" & targetRow4).Borders(xlEdgeTop).LineStyle = xlContinuous ' Borde superior
                .Range("AC" & targetRow4).Borders(xlEdgeBottom).LineStyle = xlContinuous ' Borde inferior
                .Range("AC" & targetRow4).Borders(xlEdgeBottom).Weight = xlMedium ' Borde inferior grueso
            
                .Range("AD" & targetRow4).Value = "Date"
                .Range("AD" & targetRow4 & ":AG" & targetRow4).Interior.Color = RGB(255, 255, 255) ' Fondo blanco
                .Range("AD" & targetRow4 & ":AG" & targetRow4).Merge
                .Range("AD" & targetRow4 & ":AG" & targetRow4).HorizontalAlignment = xlCenter
                .Range("AD" & targetRow4 & ":AG" & targetRow4).Font.Bold = True
                .Range("AD" & targetRow4 & ":AG" & targetRow4).Font.Size = 10
                .Range("AD" & targetRow4 & ":AG" & targetRow4).Borders(xlEdgeTop).LineStyle = xlContinuous ' Borde superior
                .Range("AD" & targetRow4 & ":AG" & targetRow4).Borders(xlEdgeBottom).LineStyle = xlContinuous ' Borde inferior
                .Range("AD" & targetRow4 & ":AG" & targetRow4).Borders(xlEdgeBottom).Weight = xlMedium ' Borde inferior grueso
            
                .Range("AH" & targetRow4).Value = "By"
                .Range("AH" & targetRow4).Interior.Color = RGB(255, 255, 255) ' Fondo blanco
                .Range("AH" & targetRow4).HorizontalAlignment = xlCenter
                .Range("AH" & targetRow4).Font.Bold = True
                .Range("AH" & targetRow4).Font.Size = 10
                .Range("AH" & targetRow4).Borders(xlEdgeTop).LineStyle = xlContinuous ' Borde superior
                .Range("AH" & targetRow4).Borders(xlEdgeBottom).LineStyle = xlContinuous ' Borde inferior
                .Range("AH" & targetRow4).Borders(xlEdgeBottom).Weight = xlMedium ' Borde inferior grueso
            
                .Range("AI" & targetRow4).Value = "Batch ID"
                .Range("AI" & targetRow4 & ":AJ" & targetRow4).Interior.Color = RGB(255, 255, 255) ' Fondo blanco
                .Range("AI" & targetRow4 & ":AJ" & targetRow4).Merge
                .Range("AI" & targetRow4 & ":AJ" & targetRow4).HorizontalAlignment = xlCenter
                .Range("AI" & targetRow4 & ":AJ" & targetRow4).Font.Bold = True
                .Range("AI" & targetRow4 & ":AJ" & targetRow4).Font.Size = 10
                .Range("AI" & targetRow4 & ":AJ" & targetRow4).Borders(xlEdgeTop).LineStyle = xlContinuous ' Borde superior
                .Range("AI" & targetRow4 & ":AJ" & targetRow4).Borders(xlEdgeBottom).LineStyle = xlContinuous ' Borde inferior
                .Range("AI" & targetRow4 & ":AJ" & targetRow4).Borders(xlEdgeBottom).Weight = xlMedium ' Borde inferior grueso
            
                .Range("AK" & targetRow4).Value = "Notes"
                .Range("AK" & targetRow4 & ":AQ" & targetRow4).Interior.Color = RGB(255, 255, 255) ' Fondo blanco
                .Range("AK" & targetRow4 & ":AQ" & targetRow4).Merge
                .Range("AK" & targetRow4 & ":AQ" & targetRow4).HorizontalAlignment = xlCenter
                .Range("AK" & targetRow4 & ":AQ" & targetRow4).Font.Bold = True
                .Range("AK" & targetRow4 & ":AQ" & targetRow4).Font.Size = 10
                .Range("AK" & targetRow4 & ":AQ" & targetRow4).Borders(xlEdgeTop).LineStyle = xlContinuous ' Borde superior
                .Range("AK" & targetRow4 & ":AQ" & targetRow4).Borders(xlEdgeBottom).LineStyle = xlContinuous ' Borde inferior
                .Range("AK" & targetRow4 & ":AQ" & targetRow4).Borders(xlEdgeBottom).Weight = xlMedium ' Borde inferior grueso
            End With
            
            extractedData = ""

            counterAnalitos = 0

            For j = 7 To 10
                If wsOrigen.Cells(i, j).Value = "Yes" Then
                    counterAnalitos = counterAnalitos + 1
                    Dim sheetName As String
                    sheetName = wsOrigen.Cells(6, j).Value
                    
                    sheetName = Trim(sheetName)
                    
                    Set hojaAnalitos = ThisWorkbook.Sheets(sheetName)
                    
                    Dim cellValue As String
                    cellValue = hojaAnalitos.Cells(7, "M").Value
                    
                    If extractedData = "" Then
                        extractedData = cellValue
                    Else
                        extractedData = extractedData & ", " & cellValue
                    End If
            
                        foundRow = hojaAnalitos.Columns("C").Find(What:=muestra, LookIn:=xlValues, LookAt:=xlWhole).Row
                    
                        targetRow5 = targetRow4 + counterAnalitos
                        
                        wsDestino.Rows(targetRow5).Insert Shift:=xlDown
                    
                        ' Colocar los títulos en la hoja de destino
                        With wsDestino
                            .Range("B" & targetRow5).Value = sheetName
                            .Range("B" & targetRow5 & ":J" & targetRow5).Merge
                            .Range("B" & targetRow5 & ":J" & targetRow5).Interior.Color = RGB(255, 255, 255)
                            .Range("B" & targetRow5 & ":J" & targetRow5).Font.Size = 9
                            .Range("B" & targetRow5 & ":J" & targetRow5).Font.Bold = False
                            .Range("B" & targetRow5 & ":J" & targetRow5).Borders.LineStyle = xlNone
                            
                            .Range("K" & targetRow5).Value = IIf(hojaAnalitos.Cells(foundRow, "H").Value < 4.99, hojaAnalitos.Cells(foundRow, "H").Value & " I", hojaAnalitos.Cells(foundRow, "H").Value)
                            .Range("K" & targetRow5 & ":S" & targetRow5).Merge
                            .Range("K" & targetRow5 & ":S" & targetRow5).Interior.Color = RGB(255, 255, 255)
                            .Range("K" & targetRow5 & ":S" & targetRow5).Font.Size = 9
                            .Range("K" & targetRow5 & ":S" & targetRow5).Font.Bold = False
                            .Range("K" & targetRow5 & ":S" & targetRow5).Borders.LineStyle = xlNone
                            
                            .Range("T" & targetRow5).Value = "mg/L"
                            .Range("T" & targetRow5 & ":U" & targetRow5).Merge
                            .Range("T" & targetRow5 & ":U" & targetRow5).Interior.Color = RGB(255, 255, 255)
                            .Range("T" & targetRow5 & ":U" & targetRow5).Font.Size = 9
                            .Range("T" & targetRow5 & ":U" & targetRow5).Font.Bold = False
                            .Range("T" & targetRow5 & ":U" & targetRow5).Borders.LineStyle = xlNone
                            
                            .Range("V" & targetRow5).Value = "1,0"
                            .Range("V" & targetRow5 & ":X" & targetRow5).Interior.Color = RGB(255, 255, 255)
                            .Range("V" & targetRow5 & ":X" & targetRow5).Merge
                            .Range("V" & targetRow5 & ":X" & targetRow5).Font.Size = 9
                            .Range("V" & targetRow5 & ":X" & targetRow5).Font.Bold = False
                            .Range("V" & targetRow5 & ":X" & targetRow5).Borders.LineStyle = xlNone
                            
                            .Range("Y" & targetRow5).Value = "2,5"
                            .Range("Y" & targetRow5 & ":AA" & targetRow5).Interior.Color = RGB(255, 255, 255)
                            .Range("Y" & targetRow5 & ":AA" & targetRow5).Merge
                            .Range("Y" & targetRow5 & ":AA" & targetRow5).Font.Size = 9
                            .Range("Y" & targetRow5 & ":AA" & targetRow5).Font.Bold = False
                            .Range("Y" & targetRow5 & ":AA" & targetRow5).Borders.LineStyle = xlNone
                            
                            .Range("AB" & targetRow5).Value = "5,0"
                            .Range("AB" & targetRow5).Interior.Color = RGB(255, 255, 255)
                            .Range("AB" & targetRow5).Font.Size = 9
                            .Range("AB" & targetRow5).Font.Bold = False
                            .Range("AB" & targetRow5).Borders.LineStyle = xlNone
                            
                            .Range("AC" & targetRow5).Value = cellValue
                            .Range("AC" & targetRow5).Interior.Color = RGB(255, 255, 255)
                            .Range("AC" & targetRow5).Font.Size = 9
                            .Range("AC" & targetRow5).Font.Bold = False
                            .Range("AC" & targetRow5).Borders.LineStyle = xlNone
                            
                            .Range("AD" & targetRow5).Value = hojaAnalitos.Cells(foundRow, "B").Value
                            .Range("AD" & targetRow5 & ":AG" & targetRow5).Interior.Color = RGB(255, 255, 255)
                            .Range("AD" & targetRow5 & ":AG" & targetRow5).Merge
                            .Range("AD" & targetRow5 & ":AG" & targetRow5).Font.Size = 9
                            .Range("AD" & targetRow5 & ":AG" & targetRow5).Font.Bold = False
                            .Range("AD" & targetRow5 & ":AG" & targetRow5).Borders.LineStyle = xlNone
                            
                            .Range("AH" & targetRow5).Value = hojaAnalitos.Cells(foundRow, "F").Value
                            .Range("AH" & targetRow5).Interior.Color = RGB(255, 255, 255)
                            .Range("AH" & targetRow5).Font.Size = 9
                            .Range("AH" & targetRow5).Font.Bold = False
                            .Range("AH" & targetRow5).Borders.LineStyle = xlNone
                            
                            .Range("AI" & targetRow5).Value = "MB" & Format(hojaAnalitos.Cells(foundRow, "B").Value, "ddmm") & Right(Format(hojaAnalitos.Cells(foundRow, "B").Value, "yyyy"), 2) & Format(hojaAnalitos.Cells(foundRow, "B").Value, "hh")
                            .Range("AI" & targetRow5 & ":AJ" & targetRow5).Interior.Color = RGB(255, 255, 255)
                            .Range("AI" & targetRow5 & ":AJ" & targetRow5).Merge
                            .Range("AI" & targetRow5 & ":AJ" & targetRow5).Font.Size = 9
                            .Range("AI" & targetRow5 & ":AJ" & targetRow5).Font.Bold = False
                            .Range("AI" & targetRow5 & ":AJ" & targetRow5).Borders.LineStyle = xlNone
                            
                            .Range("AK" & targetRow5).Value = hojaAnalitos.Cells(foundRow, "J").Value
                            .Range("AK" & targetRow5 & ":AQ" & targetRow5).Interior.Color = RGB(255, 255, 255)
                            .Range("AK" & targetRow5 & ":AQ" & targetRow5).Merge
                            .Range("AK" & targetRow5 & ":AQ" & targetRow5).Font.Size = 9
                            .Range("AK" & targetRow5 & ":AQ" & targetRow5).Font.Bold = False
                            .Range("AK" & targetRow5 & ":AQ" & targetRow5).Borders.LineStyle = xlNone
                        End With
                   End If
            Next j
            
            
            
            ' Insertar una nueva fila en la hoja de destino
            wsDestino.Rows(filaDestino).Insert Shift:=xlDown
            
            ' Condicional para asignar el valor de parentID
            If consecutivo > 99 Then
                parentID = lastLabSampleID & "-" & consecutivo
            ElseIf consecutivo > 9 Then
                parentID = lastLabSampleID & "-0" & consecutivo
            Else
                parentID = lastLabSampleID & "-00" & consecutivo
            End If
            
            wsDestino.Cells(filaDestino, "B").Value = consecutivo
            wsDestino.Cells(filaDestino, "G").Value = parentID
            wsDestino.Cells(filaDestino, "K").Value = muestra
            wsDestino.Cells(filaDestino, "Q").Value = fecha
            wsDestino.Cells(filaDestino, "U").Value = hora
            wsDestino.Cells(filaDestino, "X").Value = sampleMatrix
            wsDestino.Cells(filaDestino, "AC").Value = extractedData
            
            ' Combinar celdas en los rangos específicos
            wsDestino.Range("B" & filaDestino & ":F" & filaDestino).Merge
            wsDestino.Range("G" & filaDestino & ":J" & filaDestino).Merge
            wsDestino.Range("K" & filaDestino & ":P" & filaDestino).Merge
            wsDestino.Range("Q" & filaDestino & ":T" & filaDestino).Merge
            wsDestino.Range("U" & filaDestino & ":W" & filaDestino).Merge
            wsDestino.Range("X" & filaDestino & ":AB" & filaDestino).Merge
            wsDestino.Range("AC" & filaDestino & ":AQ" & filaDestino).Merge
            
            ' Guardar la fila actual para usarla luego
            targetRow = targetRow + 6 + counterAnalitos

            ' Incrementar el consecutivo
            consecutivo = consecutivo + 1
            
            ' Incrementar la fila de destino
            filaDestino = filaDestino + 1
        End If
    Next i
    
    wsDestino.Rows(filaAnalyticalResults.Row + 1).Delete

    ' Eliminar la fila 13
    wsDestino.Rows(13).Delete
    
    ' Reactivar alertas
    Application.DisplayAlerts = True
End Sub
