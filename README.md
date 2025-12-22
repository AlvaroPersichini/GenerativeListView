# CATIA to Excel Automation - ListView Generator

Este módulo de **VB.NET** automatiza la generación de reportes tipo "Lista de Materiales" (BOM) exportando datos desde **CATIA V5** hacia **Microsoft Excel**.

Su función principal es recorrer una estructura de producto, extraer metadatos, capturar una imagen (thumbnail) de cada parte y formatear una hoja de cálculo de Excel con esta información.

## 🚀 Características

* **Extracción de Metadatos:** Obtiene automáticamente propiedades como *Part Number*, *Description*, *Vendor Code*, *Quantity*, y *Source* (Made/Bought).
* **Parámetros de Usuario:** Lee propiedades específicas definidas por el usuario (`UserRefProperties`), como *Material* y *Material en Bruto*.
* **Generación de Miniaturas (Screenshots):**
    * Abre cada componente en una ventana nueva.
    * Cambia el fondo a blanco para una captura limpia.
    * Oculta el árbol de especificaciones y el compás.
    * Guarda una imagen `.jpg` y la inserta automáticamente en la celda de Excel correspondiente.
* **Filtrado:** Incluye lógica para omitir partes específicas (ej. archivos que comienzan con "AUX").

## 📋 Requisitos Previos

Este código requiere referencias a las librerías COM de CATIA y Excel. Asegúrate de tener referenciadas las siguientes librerías en tu proyecto:

* `INFITF` (CATIA Infrastructure)
* `ProductStructureTypeLib` (CATIA Product Structure)
* `KnowledgewareTypeLib` (CATIA Knowledge/Parameters)
* `Microsoft.Office.Interop.Excel`

### Dependencias del Proyecto
El código hace uso de un módulo auxiliar externo (no incluido en este snippet) llamado `Diccionarios` y `ExcelFormatListView`. Asegúrate de tener implementadas las siguientes funciones:
* `Diccionarios.DiccT3_Rev2(oProduct)`: Devuelve un diccionario con los productos.
* `Diccionarios.EncuentraColumna(HeaderName, Sheet)`: Devuelve la letra/índice de la columna.
* `ExcelFormatListView.FormatoListView2(Sheet)`: Aplica estilos a la hoja.

## ⚠️ Reglas de Naming Importantes

> **CRÍTICO:** Los nombres de los archivos (PartNumbers) **NO deben contener barras invertidas (`\`)**.

CATIA permite crear productos con nombres como `Cube2\Elementary Source`, pero al guardar o procesar archivos a nivel de sistema operativo, el texto antes de la barra invertida es ignorado o causa errores de ruta, lo que fallará al intentar guardar la captura de pantalla (`.jpg`).

## ⚙️ Cómo Funciona (`CompletaListView2`)

1.  **Inicialización:** Desactiva las alertas de CATIA (`DisplayFileAlerts = False`) para evitar interrupciones.
2.  **Mapeo de Columnas:** Busca dinámicamente en qué columna de Excel debe ir cada dato.
3.  **Iteración:** Recorre el diccionario de productos.
4.  **Visualización:**
    * Ejecuta `Open in New Window` para el producto actual.
    * Configura la ventana (300x300 px, fondo blanco, sin árbol).
5.  **Captura:** Toma una foto (`CaptureToFile`) en formato JPEG en el directorio especificado.
6.  **Escritura:** Vuelca los datos de texto y parámetros en las celdas de Excel.
7.  **Inserción de Imagen:** Coloca la imagen capturada dentro de la celda designada y la ajusta.
8.  **Limpieza:** Cierra la ventana temporal y restaura la configuración de visualización.

## 🛠️ Uso

```vb
' Ejemplo de llamada al procedimiento
Dim oMyProduct As ProductStructureTypeLib.Product = ... ' Tu producto raíz
Dim oMySheet As Microsoft.Office.Interop.Excel.Worksheet = ... ' Tu hoja de destino
Dim sRutaImagenes As String = "C:\Temp\ImagenesReporte"

CatiaToExcel.CompletaListView2(oMyProduct, oMySheet, sRutaImagenes)
