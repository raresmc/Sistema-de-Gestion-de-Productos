VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'@Folder("SysProd.Repositories")

'@Description("Genera numero de identificación unico")
Public Function IdMax(ByVal hoja As String, ByVal columna As String) As Long
    Dim busca As Variant
    Dim rango As Range
    
    Set rango = ThisWorkbook.Sheets(hoja).Range(columna) ' Columna ID
    Set busca = rango.Find(Application.WorksheetFunction.Max(rango), _
    LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not busca Is Nothing Then
        IdMax = busca.value + 1
    Else
        IdMax = 1
    End If
End Function

'@Description("Retorna el numero de registros duplicados")
'Public Function ValidarDuplicado( _
'    ByVal hoja As String, _
'    ByVal columna As String, _
'    ByVal nombreProd _
'    ) As Integer
'
'    Sheets(hoja).Activate
'
'    Dim cuenta As Integer
'    cuenta = Application.CountIf(Range(columna), UCase(nombreProd))
'
'    ' Si el registro existe retornamos el numero de registros duplicados
'    If cuenta > 0 Then
'        MsgBox "Registro ya registrado"
'        ValidarDuplicado = cuenta
'    Else
'        ValidarDuplicado = 0
'    End If
'End Function


