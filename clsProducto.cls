VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("SysProd.Entidades")

Private Type TProducto
    Identificador As Long
    Nombre As String
    Imagen As String
    Estado As String
End Type
Private this As TProducto
    
Implements IProducto

Private Property Get IProducto_Identificador() As Long
    IProducto_Identificador = this.Identificador
End Property


