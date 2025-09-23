VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ISolicitudService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Function CreateSolicitud(ByVal paraExpediente As EExpediente) As ESolicitud
End Function

' Guarda una solicitud existente
' @param solicitud: Objeto ESolicitud a guardar
' @return: ID de la solicitud guardada (Long), 0 en caso de error
Public Function SaveSolicitud(ByVal solicitud As ESolicitud) As Long
End Function

' Obtiene una solicitud por su ID
' @param solicitudId: ID de la solicitud a obtener
' @return: Objeto ESolicitud encontrado
Public Function ObtenerSolicitudPorId(ByVal SolicitudID As Long) As ESolicitud
End Function

' Cambia el estado de una solicitud validando la transici?n
' @param solicitudId: ID de la solicitud
' @param nuevoEstadoId: ID del nuevo estado
' @return: True si el cambio fue exitoso, False en caso contrario
Public Function CambiarEstadoSolicitud(ByVal SolicitudID As Long, ByVal nuevoEstadoId As Long) As Boolean
End Function
