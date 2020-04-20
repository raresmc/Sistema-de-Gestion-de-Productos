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

Private Property Get IProducto_Nombre() As String
    IProducto_Nombre = this.Nombre
End Property

Private Property Get IProducto_Imagen() As String
    IProducto_Imagen = this.Imagen
End Property

Private Property Get IProducto_Estado() As String
    IProducto_Estado = this.Estado
End Property

' Method Factory
Public Sub FillData( _
    ByRef Identificador As Long, _
    ByRef Nombre As String, _
    ByRef Imagen As String, _
    ByRef Estado As String _
    )
    If Len(this.Nombre) = 0 Then
        this.Identificador = Identificador
        this.Nombre = Nombre
        this.Imagen = Imagen
        this.Estado = Estado
    End If
End Sub

