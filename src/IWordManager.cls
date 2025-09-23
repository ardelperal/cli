VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWordManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Function AbrirDocumento(ByVal rutaDocumento As String) As Boolean
End Function

' Establece el texto de un marcador (bookmark) en el documento
' Par?metros:
'   - BookmarkName: Nombre del marcador
'   - Value: Valor a establecer en el marcador
' Retorna: True si se estableci? correctamente, False en caso contrario
Public Function SetBookmarkText(ByVal BookmarkName As String, ByVal value As String) As Boolean
End Function

' Obtiene el texto de un marcador (bookmark) del documento
' Par?metros:
'   - BookmarkName: Nombre del marcador
' Retorna: El texto del marcador o cadena vac?a si no existe
Public Function GetBookmarkText(ByVal BookmarkName As String) As String
End Function

' Guarda el documento actualmente abierto
' Retorna: True si se guard? correctamente, False en caso contrario
Public Function GuardarDocumento() As Boolean
End Function
Public Sub Dispose()
End Sub
