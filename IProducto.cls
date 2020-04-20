VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("SysProd.Contracts")
'@Interface
Option Explicit

Const HojaTrabajo As String = "Productos"
Const columna As String = "A:A"

Public Property Get Identificador() As Long
    Dim Codigo As clsMaster
    Set Codigo = New clsMaster
    
     Identificador = Codigo.IdMax(HojaTrabajo, columna)
End Property

Public Property Get Nombre() As String
    
End Property

Public Property Get Imagen() As String

End Property

Public Property Get Estado() As String

End Property
