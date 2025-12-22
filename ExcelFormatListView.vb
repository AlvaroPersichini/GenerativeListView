Module ExcelFormatListView

    ' Escribir el prospecto de ésta función.-
    Private Function FindLastRowWithData(oWorkSheet As Microsoft.Office.Interop.Excel.Worksheet) As Integer
        Dim oRange As Microsoft.Office.Interop.Excel.Range = oWorkSheet.Range("A1", "U1") ' solo trabaja desde columna A hasta U
        Dim i As Integer
        Dim j As Integer = 3
        Dim strColName As String
        For Each c As Microsoft.Office.Interop.Excel.Range In oRange.Columns
            strColName = Left(c.Columns.Address(False, False, ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1), 1)
            i = oWorkSheet.Range(strColName & oWorkSheet.Rows.Count).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
            If i > j Then
                j = i
            End If
        Next
        Return j
    End Function


    '//************************************************************************************************************
    '// Da formato a una hoja de excel para ser completada con la información del procedimiento "CompletaLIstView2"
    '// Falta mejorar las opciones de las imagenes
    '//************************************************************************************************************
    Sub FormatoListView2(oWorkSheetListView As Microsoft.Office.Interop.Excel.Worksheet)

        oWorkSheetListView.Activate() : oWorkSheetListView.Name = "ListView"
        Dim i As Integer = FindLastRowWithData(oWorkSheetListView)
        Dim oWorkBook As Microsoft.Office.Interop.Excel.Workbook = oWorkSheetListView.Parent

        'Está asignando el item 1 del total de todas las ventanas.
        Dim viewListView As Microsoft.Office.Interop.Excel.WorksheetView = oWorkBook.Windows.Item(1).SheetViews.Item(1) : viewListView.DisplayGridlines = False
        Dim oRangoEncabezado As Microsoft.Office.Interop.Excel.Range = oWorkSheetListView.Range("A1", "U2")
        Dim oRangoCuerpo As Microsoft.Office.Interop.Excel.Range = oWorkSheetListView.Range("A3", "U3")
        Dim strColumnLetter As String
        Dim oCurrentRange As Microsoft.Office.Interop.Excel.Range
        Dim a As String
        Dim b As String


        ' // Arma el diccionario con los textos del encabezado. Si a futuro se requieren otras columnas
        ' // hay que modificar esto. Se pueden armar diccionarios o listas aparte y luego pasarlas como argumentos
        Dim oDicListViewColumnText As New Dictionary(Of String, String) From {
            {"A1", "Grupo"},
            {"B1", "Prefix"},
            {"C1", "Part Number"},
            {"D1", "Description"},
            {"E1", "Cantidad"},
            {"F1", "Conjunto - Parte"},
            {"G1", "Made or Bought"},
            {"H1", "-Libre-"},
            {"I1", "Vendor_Code_ID"},
            {"J1", "-libre-"},
            {"K1", "Material en Bruto"},
            {"L1", "Material"},
            {"M1", "Terminacion Superf"},
            {"N1", "Tratamiento Termico"},
            {"O1", "Peso"},
            {"P1", "Costo Unitario Estimado"},
            {"Q1", "Supplier/Vendor"},
            {"R1", "Lead Time [Week]"},
            {"S1", "Documento"},
            {"T1", "Obs."},
            {"U1", "Image"}
        }
        For Each kvp As KeyValuePair(Of String, String) In oDicListViewColumnText
            oWorkSheetListView.Range(kvp.Key).Value = kvp.Value
        Next

        ' // Bordes del encabezado
        For Each c As Microsoft.Office.Interop.Excel.Range In oRangoEncabezado.Cells
            With c
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeBottom).LineStyle = -4119
            End With
        Next

        ' // Fuente, tamaño y alineado de todo el documento
        With oWorkSheetListView.Cells
            .Font.Name = "Tahoma"
            .Font.Size = 8
            .HorizontalAlignment = -4108
            .VerticalAlignment = -4107
        End With

        ' // Fuente, tamaño y alineado del encabezado
        With oWorkSheetListView
            .Range("A1", "U1").Orientation = 90
            .Range("A1", "U1").Font.Bold = True
            .Range("A1", "C1").Interior.Color = RGB(204, 255, 255)
            .Range("D1", "U1").Interior.ColorIndex = 15
            .Range("A2", "U2").Interior.ColorIndex = 15
        End With


        ' Hace AutoFit pero a la columna de imagenes no.
        ' Aca hay que incluir la opcion de que si la planilla va a tener imagenes entonces que no haga AutoFit,
        ' pero si son incluidas las imagenes, no debería hacer autofit.
        For Each C As Microsoft.Office.Interop.Excel.Range In oRangoEncabezado
            C.EntireColumn.AutoFit()
        Next


        ' // Formato aplicado a todo el cuerpo
        With oWorkSheetListView
            .Range("A3", "U" & i).VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter
            .Range("A3", "U" & i).HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
            .Range("U3", "U" & i).RowHeight = 100
            .Range("U3", "U" & i).ColumnWidth = 18
        End With


        ' Para aplicar los bordes a cada columna hasta la última fila de datos,
        ' hay que hacer estos pasos para armar el rango
        For Each c As Microsoft.Office.Interop.Excel.Range In oRangoCuerpo
            strColumnLetter = Left(c.Address(False, False, ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1), 1)
            a = c.Address(False, False, ReferenceStyle:=Microsoft.Office.Interop.Excel.XlReferenceStyle.xlA1)
            b = strColumnLetter & FindLastRowWithData(oWorkSheetListView)
            oCurrentRange = oWorkSheetListView.Range(a, b)
            With oCurrentRange
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeRight).LineStyle = 1
                .Borders().Item(Microsoft.Office.Interop.Excel.XlBordersIndex.xlEdgeLeft).LineStyle = 1
            End With
        Next

    End Sub









End Module
