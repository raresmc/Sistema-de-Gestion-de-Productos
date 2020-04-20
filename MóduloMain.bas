Attribute VB_Name = "MóduloMain"
Option Explicit
'@Folder("SysProd")
Public Sub Main()
    Dim objProd As IProducto
    Set objProd = ProductoFactory.Create("Producto1", "Ruta imagen", "Activo")
    
    Debug.Print objProd.Identificador, objProd.Nombre, objProd.Imagen, objProd.Estado
    
End Sub
