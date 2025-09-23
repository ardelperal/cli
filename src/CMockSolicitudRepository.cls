VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockSolicitudRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements ISolicitudRepository

' --- Variables de estado para configurar el comportamiento del mock ---
Private m_SaveSolicitud_Result As Long
Private m_ObtenerSolicitudPorId_Result As ESolicitud
Private m_SaveSolicitud_LastSolicitud As ESolicitud

' --- Propiedades de seguimiento para las aserciones ---
Public SaveSolicitud_CallCount As Long
Public LastSavedSolicitud As ESolicitud
Public Update_LastUsuarioModificacion As String

' Constructor
Private Sub Class_Initialize()
    Reset
End Sub

' --- Propiedades de conveniencia para verificaci?n booleana ---
Public Property Get SaveSolicitudCalled() As Boolean
    SaveSolicitudCalled = (SaveSolicitud_CallCount > 0)
End Property

' --- Propiedad esp?a para inspeccionar datos pasados al mock ---
Public Property Get SaveSolicitud_LastSolicitud() As ESolicitud
    Set SaveSolicitud_LastSolicitud = m_SaveSolicitud_LastSolicitud
End Property

' --- M?todos p?blicos para configurar el mock desde los tests ---
Public Sub ConfigureSaveSolicitud(ByVal returnValue As Long)
    m_SaveSolicitud_Result = returnValue
End Sub

Public Sub ConfigureObtenerSolicitudPorId(ByVal solicitud As ESolicitud)
    Set m_ObtenerSolicitudPorId_Result = solicitud
End Sub

' --- Resetea el estado del mock para aislar los tests ---
Public Sub Reset()
    m_SaveSolicitud_Result = 0
    Set m_ObtenerSolicitudPorId_Result = Nothing
    SaveSolicitud_CallCount = 0
    Set LastSavedSolicitud = Nothing
    Update_LastUsuarioModificacion = ""
    Set m_SaveSolicitud_LastSolicitud = Nothing
End Sub



Private Function ISolicitudRepository_SaveSolicitud(ByVal solicitud As ESolicitud) As Long
    Set m_SaveSolicitud_LastSolicitud = solicitud
    SaveSolicitud_CallCount = SaveSolicitud_CallCount + 1
    Set LastSavedSolicitud = solicitud
    Update_LastUsuarioModificacion = solicitud.usuarioModificacion
    ISolicitudRepository_SaveSolicitud = m_SaveSolicitud_Result
End Function

' ============================================================================
' M?TODOS P?BLICOS DE CONVENIENCIA
' ============================================================================



' M?todo p?blico de conveniencia para SaveSolicitud
Public Function SaveSolicitud(ByVal solicitud As ESolicitud) As Long
    SaveSolicitud = ISolicitudRepository_SaveSolicitud(solicitud)
End Function

' --- Implementaci?n de la interfaz para ObtenerSolicitudPorId ---
Private Function ISolicitudRepository_ObtenerSolicitudPorId(ByVal SolicitudID As Long) As ESolicitud
    Set ISolicitudRepository_ObtenerSolicitudPorId = m_ObtenerSolicitudPorId_Result
End Function

' M?todo p?blico de conveniencia para ObtenerSolicitudPorId
Public Function ObtenerSolicitudPorId(ByVal SolicitudID As Long) As ESolicitud
    Set ObtenerSolicitudPorId = ISolicitudRepository_ObtenerSolicitudPorId(SolicitudID)
End Function
