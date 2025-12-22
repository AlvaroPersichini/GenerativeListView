Option Explicit On


Module CatiaToExcel



    ' **************************************************************************************************************************
    ' IMPORTANTE !!!!!!!!!!!!!!!!!!!!!!!!
    ' Habria que hacer un procedimiento para comprobar que no hay "\" (backlash) en los nombres de los archivos!
    ' Importante: Naming Rules:
    ' Create a New product and give it a name with a backslash: "Cube2\Elementary Source".
    ' Save this product And you will see that all the words before the backslash, And the backslash, are Not taken into account.

    ' Completa LIstView
    Sub CompletaListView2(oProduct As ProductStructureTypeLib.Product,
                          oSheetListView As Microsoft.Office.Interop.Excel.Worksheet,
                          strDir As String)


        ' Antes de empezar a completar da formato a la hoja
        ExcelFormatListView.FormatoListView2(oSheetListView)


        ' indicar que es esto
        Dim sFullPathFileName As String


        ' Se referencia la aplicación a partir del oProduct que se recibe por parámetro.
        ' Otra opción puede ser, recibir por parámetro la nueva ventana que fue abierta.
        Dim oAppCATIA As INFITF.Application = oProduct.Application


        ' Bloquea la interacccion para evitar cambio de "ActiveDocument" por un click sobre CATIA.
        ' Si bien ya hay que desactivar la interaccion antes de entrar a este método, lo vuelvo a hacer acá.
        ' Tambien deshabilita los mensajes emergentes.
        ' Hay un problema con la función "Interactive", porque al terminar la ejecución, CATIA queda con menu greyed-out
        ' oAppCATIA.Interactive = False
        oAppCATIA.DisplayFileAlerts = False


        ' Este objeto selection va a servir para hacer un "OpenInNewWindow" por cada product para poder sacar las capturas de pantalla
        ' Ver que apenas creo el objeto, luego hago "clear"
        Dim oSelection As INFITF.Selection = oAppCATIA.ActiveDocument.Selection : oSelection.Clear()


        ' Indicar que es esto
        Dim oShape As Microsoft.Office.Interop.Excel.Shape


        ' Esta parte de parametros no esta bien pulida
        Dim oUserRefParameters As KnowledgewareTypeLib.Parameters

        'Para completar la pestaña "ListView" se necesita un diccionario de tipo 3
        Dim oDiccType3 As Dictionary(Of String, PwrProduct) = Diccionarios.DiccT3_Rev2(oProduct)


        ' Va a buscar a Excel los nombres de las columnas y los referencia a una variable

        Dim ListViewPartNumCol As String = Diccionarios.EncuentraColumna("Part Number", oSheetListView)
        Dim ListViewDescriptionCol As String = Diccionarios.EncuentraColumna("Description", oSheetListView)
        Dim ListViewVendor_Code_IDCol As String = Diccionarios.EncuentraColumna("Vendor_Code_ID", oSheetListView)
        Dim ListViewMaterialenBrutoCol As String = Diccionarios.EncuentraColumna("Material en Bruto", oSheetListView)
        Dim ListViewTrataTermicoCol As String = Diccionarios.EncuentraColumna("Tratamiento Termico", oSheetListView)
        Dim ListViewCantidad As String = Diccionarios.EncuentraColumna("Cantidad", oSheetListView)
        Dim ListViewConjuntoParteCol As String = Diccionarios.EncuentraColumna("Conjunto - Parte", oSheetListView)
        Dim ListViewMadeOrBoughtCol As String = Diccionarios.EncuentraColumna("Made or Bought", oSheetListView)
        Dim ListViewMaterialCol As String = Diccionarios.EncuentraColumna("Material", oSheetListView)
        Dim ListViewImageCol As String = Diccionarios.EncuentraColumna("Image", oSheetListView)


        'Completa el ListView:
        Dim i As Integer = 3 'Esta linea es para ir hacia abajo en cada linea de excel

        For Each kvp As KeyValuePair(Of String, PwrProduct) In oDiccType3

            ' En primera instancia, se evalúa que se quiere computar.
            ' Si algun tipo de product se quiere filtar, se puede utilizar "Continue for" en esta primera parte.
            ' Todo lo que se incluya luego de este procedimiento, se computa.
            ' Por ejemplo No computar los "AUX" (o algun otro nombre que se incluya acá)
            If Left(kvp.Value.Product.PartNumber, 3) = "AUX" Then
                Continue For
            End If

            ' Agrego a la seleccion el product que voy a hacer "OpenInNewWindow"
            oSelection.Add(kvp.Value.Product)


            ' Abro la ventana para sacar la captura
            oAppCATIA.StartCommand("Open in New Window")


            oAppCATIA.RefreshDisplay = True ' ésto no tengo claro si sirve para algo.
            oSelection.Clear()

            ' *****************************************************
            ' Sobre la Ventana abierta se realiza todo lo siguiente
            ' *****************************************************

            ' oCurrentWindow: esto esta asegurado porque al ingresar a este metodo antes se ejecutó: " oAppCATIA.Interactive = False"
            ' Pero creo que se podría mejorar si se utiliza Windows.Item("Nombre de la ventana que corresponde al product que ingresó por parámetro"):
            ' y luego se activa esa ventana. Aunque para hacer esto, se necesita ver que tipo de doc se va a abrir en la nueva ventana.
            Dim oCurrentWindow As INFITF.Window = oAppCATIA.ActiveWindow
            Dim oSpecsAndGeomWindow As INFITF.SpecsAndGeomWindow = oCurrentWindow '(QueryInterface) (Pag.235)

            ' El oViewer3D antes estaba resuelto con la línea: "Dim objViewer3D As INFITF.Viewer3D = objAppCATIA.ActiveWindow.ActiveViewer"
            Dim oViewer3D As INFITF.Viewer3D = oSpecsAndGeomWindow.Viewers.Item(1) '(QueryInterface) (Pag.235)
            Dim oViewPoint3D As INFITF.Viewpoint3D

            ' En las cámaras pasa lo mismo que en "oCurrentWindow", ya que para poder referenciar las camaras se necesita el documento activo
            ' Trabajar con la sentencia "ActiveDocuemnt" o "ActiveWindow" en un método que recibe por parámetro el oProduct tiene cierta inconsistencia,
            ' ya que debería estár todo referenciado a el parámetro que ha ingresado como argumento. De todas formas, se utiliza el " oAppCATIA.Interactive = False"
            ' para que el usuario no cambie el documento o ventana activa.
            Dim oCameras As INFITF.Cameras = oAppCATIA.ActiveDocument.Cameras
            Dim oCamera3D As INFITF.Camera3D = oCameras.Item(1)  ' las primeras 7 camaras son de tipo "Camera3D". Camara 1 = isometrica  '(QueryInterface) (Pag.235)


            ' Estos arrays son para el color de fondo de la ventana. Estos arreglos contiene tipo genérico "Object", ya que en la documentación de CATIA
            ' se indica que deben ser del tipo Variant
            Dim arrBackgroundColor(2) As Object
            Dim arrWhiteColor(2) As Object
            arrWhiteColor(0) = 1
            arrWhiteColor(1) = 1
            arrWhiteColor(2) = 1

            ' ***************************************************************************
            ' Seteo de la ventana para poder tomar la captura de pantalla
            ' ***************************************************************************
            oSpecsAndGeomWindow.Layout = INFITF.CatSpecsAndGeomWindowLayout.catWindowGeomOnly  ' Apaga el arbol de especificaciones
            oViewPoint3D = oCamera3D.Viewpoint3D   ' Setteo de la vista en la que se va a tomar la captura
            oAppCATIA.StartCommand("Compass")  ' Oculta el compass
            oViewer3D.GetBackgroundColor(arrBackgroundColor)  ' Toma el color actual de fondo y lo almacena en "arrBackgroundColor" para luego reestablecerlo.
            oViewer3D.PutBackgroundColor(arrWhiteColor)  ' Luego, setea el fondo de color a blanco para tomar la captura
            oCurrentWindow.Height = 300  ' Altura de la pantalla para la captura
            oCurrentWindow.Width = 300   ' Ancho de la pantalla para la captura
            oViewer3D.Update()
            oAppCATIA.RefreshDisplay = True
            oViewer3D.Reframe()


            ' The full pathname of the file into which you want to store the captured image 
            sFullPathFileName = strDir & "\" & kvp.Value.Product.PartNumber & ".jpg"



            'Antes de llenar las celdas con valores, se le debe dar formato a la columna de Part Number
            oSheetListView.Range("C3", "C" & oDiccType3.Count).NumberFormat = "@"

            With oSheetListView
                .Cells(i, ListViewPartNumCol) = kvp.Value.Product.PartNumber
                .Cells(i, ListViewDescriptionCol) = kvp.Value.Product.DescriptionRef
                .Cells(i, ListViewVendor_Code_IDCol) = kvp.Value.Product.Definition ' Vendor_Code_ID
                .Cells(i, ListViewMadeOrBoughtCol) = Left(kvp.Value.Source, 1)
                .Cells(i, ListViewCantidad) = kvp.Value.Quantity
                .Cells(i, ListViewConjuntoParteCol) = kvp.Value.ProductType
            End With


            ' ****************************************************************************
            ' Propiedades de Usuario
            ' (1) Material en Bruto
            ' (2) Material
            ' NOTA: Si se agregan propiedades de usuario, entonces hay que agregarlas acá
            ' ****************************************************************************
            oUserRefParameters = kvp.Value.Product.ReferenceProduct.UserRefProperties

            If oUserRefParameters.Count <> 0 Then 'Primero ver si hay propiedades de usuario
                For Each Parametro As KnowledgewareTypeLib.Parameter In oUserRefParameters  'Luego, buscar cuales hay
                    If Parametro.Name = "Material" Then
                        oSheetListView.Cells(i, ListViewMaterialCol) = oUserRefParameters.Item("Material").ValueAsString()
                    End If
                    If Parametro.Name = "Material en Bruto" Then
                        oSheetListView.Cells(i, ListViewMaterialenBrutoCol) = oUserRefParameters.Item("Material en Bruto").ValueAsString()
                    End If
                Next
            End If


            ' Toma la captura y la amlmacena en el directorio indicado
            oViewer3D.CaptureToFile(INFITF.CatCaptureFormat.catCaptureFormatJPEG, sFullPathFileName)


            ' Reset antes de cerrar la ventana!
            oViewer3D.PutBackgroundColor(arrBackgroundColor)
            oAppCATIA.StartCommand("Compass")
            oSpecsAndGeomWindow.Layout = INFITF.CatSpecsAndGeomWindowLayout.catWindowSpecsAndGeom
            oViewer3D.Reframe()
            oCurrentWindow.WindowState = INFITF.CatWindowState.catWindowStateNormal
            oCurrentWindow.WindowState = INFITF.CatWindowState.catWindowStateMaximized


            ' Una vez que ya he tomado la captura de pantalla cierro la ventana activa
            ' Hay un condicional para no cerrar el producto raíz
            ' No se si es la mejor manera, hay que ver si existe otra forma mejor.
            ' lo que hace esta forma es, ver que tipo de padre tiene el product.
            ' El product raíz va a tener de padre un objeto de tipo collection (products)
            ' Creo que a lo que me refería es a que el Raiz no va a tener un padre de tipo "collection".
            ' El padre del root que es?
            If TypeName(kvp.Value.Product.Parent) = "Products" Then
                kvp.Value.Product.Application.ActiveDocument.Close()
            End If


            ' Esto es para ubicar la imagen en la celda 
            Dim cl As Microsoft.Office.Interop.Excel.Range = oSheetListView.Cells(i, ListViewImageCol)
            Dim clLeft As Single = cl.Left
            Dim clTop As Single = cl.Top


            ' **************************************************************************************************************************
            ' IMPORTANTE !!!!!!!!!!!!!!!!!!!!!!!!
            ' Habria que hacer un procedimiento para comprobar que no hay "backlash" en los nombres de los archivos!
            ' "Valvula 5_2 24V.CATPart": este archivo tenía backlash y daba error
            ' Importante: Naming Rules:
            ' Create a New product and give it a name with a backslash: "Cube2\Elementary Source".
            ' Save this product And you will see that all the words before the backslash, And the backslash, are Not taken into account.

            If FileIO.FileSystem.FileExists(sFullPathFileName) Then
                oShape = oSheetListView.Shapes.AddPicture(sFullPathFileName, False, True, clLeft + 5.5, clTop + 5, 80, 80)
            End If

            ' Siguiente línea de la hoja de trabajo
            i += 1
        Next

        ' Una vez completado con todos los datos, se da formato nuevamente
        ExcelFormatListView.FormatoListView2(oSheetListView)

    End Sub













End Module
