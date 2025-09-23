VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWorkflowService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Function ValidateTransition(ByVal SolicitudID As Long, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Boolean
End Function

' Obtiene los siguientes estados v?lidos para el flujo refactorizado
Public Function GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Object
End Function

' Verifica si un estado es final en el nuevo flujo (Aprobada = ID 7)
Public Function IsEstadoFinal(ByVal estadoActual As String) As Boolean
End Function
