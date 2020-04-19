Attribute VB_Name = "MóduloMain"
Option Explicit
'@Folder("SysProd")
Sub Main()
    Dim cod As IProducto
    Set cod = New IProducto
    
    Dim nuevoId As Long
    nuevoId = cod.Identificador
    Debug.Print nuevoId
End Sub
