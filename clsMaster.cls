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

Public Function IdMax(ByVal hoja As String, ByVal Columna As String) As Long
    Dim busca As Variant
    Dim rango As Range
    
    Set rango = ThisWorkbook.Sheets(hoja).Range(Columna) ' Columna ID
    Set busca = rango.Find(Application.WorksheetFunction.Max(rango), _
    LookIn:=xlValues, LookAt:=xlWhole)
    
    If Not busca Is Nothing Then
        IdMax = busca.value + 1
    Else
        IdMax = 1
    End If

End Function


