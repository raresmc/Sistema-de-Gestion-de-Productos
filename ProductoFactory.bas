Attribute VB_Name = "ProductoFactory"
'@Folder("SysProd.Repositories")
'@Description("Factory")
Option Explicit

Public Function Create( _
    ByVal Nombre As String, _
    ByVal Imagen As String, _
    ByVal Estado As String _
    ) As IProducto
    Dim NewProduct As clsProducto
    Set NewProduct = New clsProducto
    
    Dim NewId As IProducto
    Set NewId = New IProducto
    
    
    NewProduct.FillData NewId.Identificador, Nombre, Imagen, Estado
    Set Create = NewProduct
End Function

