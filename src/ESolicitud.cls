VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ESolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private m_idSolicitud As Long
Private m_idExpediente As Long
Private m_tipoSolicitud As String
Private m_subTipoSolicitud As String
Private m_codigoSolicitud As String
Private m_idEstadoInterno As Long ' REFACTORIZADO: De String a Long
Private m_fechaCreacion As Date
Private m_usuarioCreacion As String
Private m_fechaPaseTecnico As Date
Private m_fechaCompletadoTecnico As Date
Private m_fechaModificacion As Date ' A?ADIDO: Campo de auditor?a
Private m_usuarioModificacion As String ' A?ADIDO: Campo de auditor?a
Private m_observaciones As String

' Propiedades P?blicas
Public Property Get idSolicitud() As Long: idSolicitud = m_idSolicitud: End Property
Public Property Let idSolicitud(ByVal value As Long): m_idSolicitud = value: End Property
Public Property Get idExpediente() As Long: idExpediente = m_idExpediente: End Property
Public Property Let idExpediente(ByVal value As Long): m_idExpediente = value: End Property
Public Property Get tipoSolicitud() As String: tipoSolicitud = m_tipoSolicitud: End Property
Public Property Let tipoSolicitud(ByVal value As String): m_tipoSolicitud = value: End Property
Public Property Get subTipoSolicitud() As String: subTipoSolicitud = m_subTipoSolicitud: End Property
Public Property Let subTipoSolicitud(ByVal value As String): m_subTipoSolicitud = value: End Property
Public Property Get codigoSolicitud() As String: codigoSolicitud = m_codigoSolicitud: End Property
Public Property Let codigoSolicitud(ByVal value As String): m_codigoSolicitud = value: End Property
Public Property Get idEstadoInterno() As Long: idEstadoInterno = m_idEstadoInterno: End Property
Public Property Let idEstadoInterno(ByVal value As Long): m_idEstadoInterno = value: End Property
Public Property Get fechaCreacion() As Date: fechaCreacion = m_fechaCreacion: End Property
Public Property Let fechaCreacion(ByVal value As Date): m_fechaCreacion = value: End Property
Public Property Get usuarioCreacion() As String: usuarioCreacion = m_usuarioCreacion: End Property
Public Property Let usuarioCreacion(ByVal value As String): m_usuarioCreacion = value: End Property
Public Property Get fechaPaseTecnico() As Date: fechaPaseTecnico = m_fechaPaseTecnico: End Property
Public Property Let fechaPaseTecnico(ByVal value As Date): m_fechaPaseTecnico = value: End Property
Public Property Get fechaCompletadoTecnico() As Date: fechaCompletadoTecnico = m_fechaCompletadoTecnico: End Property
Public Property Let fechaCompletadoTecnico(ByVal value As Date): m_fechaCompletadoTecnico = value: End Property
Public Property Get fechaModificacion() As Date: fechaModificacion = m_fechaModificacion: End Property
Public Property Let fechaModificacion(ByVal value As Date): m_fechaModificacion = value: End Property
Public Property Get usuarioModificacion() As String: usuarioModificacion = m_usuarioModificacion: End Property
Public Property Let usuarioModificacion(ByVal value As String): m_usuarioModificacion = value: End Property
Public Property Get observaciones() As String: observaciones = m_observaciones: End Property
Public Property Let observaciones(ByVal value As String): m_observaciones = value: End Property
