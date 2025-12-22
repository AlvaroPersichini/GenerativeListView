Module Module1

    Sub Main()




        ' ***************************************** ListView2  ***********************
        Dim myExcel As Microsoft.Office.Interop.Excel.Application = CType(GetObject(, "Excel.Application"), Microsoft.Office.Interop.Excel.Application) ' Enlaza EXCEL y comprobar estado. (pag.177)
        Dim oWorkbook As Microsoft.Office.Interop.Excel.Workbook = myExcel.ActiveWorkbook
        Dim oWorkSheet As Microsoft.Office.Interop.Excel.Worksheet = oWorkbook.Worksheets.Item(1)

        Dim oCATIA As New CatiaSession   ' Instancia un objeto de clase CatiaSession para enlazar la aplicacion y comprobar el estado actual de la misma. (pag.177)
        ' esta parte hay que seguirlaporque supone que CATIA esta abierta y con un documento de producto activo. hay que hacer un manejador de esto.
        Dim oProductDocument As ProductStructureTypeLib.ProductDocument = oCATIA.Application.ActiveDocument
        Dim oPorduct As ProductStructureTypeLib.Product = oProductDocument.Product

        oCATIA.Application.DisplayFileAlerts = False
        myExcel.ScreenUpdating = False

        ' oCATIA.AppCATIA.Interactive = False ' Esto hay que verlo porque CATIA queda sin menues luego de correr este programa con esta linea.

        To_Excel.CompletaListView2(oPorduct, oWorkSheet, "C:\Temp")
        Diccionarios.FormatoListView2(oWorkSheet)

        oCATIA.AppCATIA.Interactive = True
        myExcel.ScreenUpdating = True






    End Sub

End Module
