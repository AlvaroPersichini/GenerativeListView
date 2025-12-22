Module Diccionarios




    ' Se ingresa con el string de la columna y sale con la letra de esa columna
    ' Letra que representa la columna: A - B - C - D...
    ' Los nombres de los parámetros son buscados desde la columna "A" a las "Z"
    ' NOTA: Si hay parametros mas allá de la Z no seran encontrados
    ' Asigna la letra de la columna encontrada a una variable de tipo string
    Public Function EncuentraColumna(strTituloColumna As String, oWorkSheet As Microsoft.Office.Interop.Excel.Worksheet) As String
        Dim myExcel As Microsoft.Office.Interop.Excel.Application = CType(GetObject(, "Excel.Application"), Microsoft.Office.Interop.Excel.Application)
        Dim oWorkbook As Microsoft.Office.Interop.Excel.Workbook = myExcel.ActiveWorkbook
        Dim oRangeContainingCell As Microsoft.Office.Interop.Excel.Range = oWorkSheet.Range("A1:Z1").Find(strTituloColumna, , , Microsoft.Office.Interop.Excel.XlLookAt.xlWhole)

        If oRangeContainingCell Is Nothing Then
            MsgBox("no se ha encontrado la columna")
        End If

        Dim strResultColum As String = Left(oRangeContainingCell.Address(RowAbsolute:=False, ColumnAbsolute:=False), 1)
        Return strResultColum
    End Function





    Public Function DiccT3_Rev2(oRootProduct As ProductStructureTypeLib.Product) As Dictionary(Of String, PwrProduct)
        Static Dim oDictionary As New Dictionary(Of String, PwrProduct)
        Static Dim intCont As Integer = 1
        Dim strProductType As String

        ' Para incluir el rootProduct lo resolví de esta manera, utilizando un contador.
        If intCont = 1 Then
            Dim PPRoot As New PwrProduct
            With PPRoot
                .Product = oRootProduct
                .Quantity = 1
                .ProductType = Replace(TypeName(oRootProduct.ReferenceProduct.Parent), "ProductDocument", "C")
                .Source = Replace([Enum].GetName(GetType(ProductStructureTypeLib.CatProductSource), oRootProduct.Source), "catProduct", "")
            End With
            oDictionary.Add(oRootProduct.PartNumber, PPRoot)
        End If
        intCont += 1
        For Each Product As ProductStructureTypeLib.Product In oRootProduct.Products
            Dim PP As New PwrProduct
            With PP
                strProductType = Replace(TypeName(Product.ReferenceProduct.Parent), "ProductDocument", "C")
                strProductType = Replace(strProductType, "PartDocument", "P")
                .Product = Product
                .ProductType = strProductType
                .Source = Replace([Enum].GetName(GetType(ProductStructureTypeLib.CatProductSource), Product.Source), "catProduct", "")
            End With
            If oDictionary.ContainsKey(PP.Product.PartNumber) Then
                oDictionary.Item(PP.Product.PartNumber).Quantity = oDictionary.Item(PP.Product.PartNumber).Quantity + 1
                If oRootProduct.Products.Count > 0 Then
                    DiccT3_Rev2(PP.Product)
                End If
                GoTo Finish
            Else
                PP.Quantity = 1
                oDictionary.Add(PP.Product.PartNumber, PP)
                If oRootProduct.Products.Count > 0 Then
                    DiccT3_Rev2(PP.Product)
                End If
            End If
Finish:
        Next
        Return oDictionary
    End Function










End Module
