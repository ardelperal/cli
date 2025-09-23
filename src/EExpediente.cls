VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EExpediente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private m_idExpediente As Long
Private m_Nemotecnico As String
Private m_Titulo As String
Private m_ResponsableCalidad As String
Private m_ResponsableTecnico As String
Private m_Pecal As String
Private m_NumeroExpediente As String
Private m_EstadoExpediente As String
Private m_EmailResponsable As String
Private m_JefeProyecto As String
Private m_ContratistaPrincipal As String
Private m_FechaInicio As Date
Private m_FechaFinPrevista As Date
Private m_FechaFinContrato As Date
Private m_FechaFinGarantia As Date
Private m_Descripcion As String
Private m_fechaCreacion As Date
Private m_Estado As String
Private m_IdUsuarioCreador As Long
Private m_NombreUsuarioCreador As String

'******************************************************************************
' PROPIEDADES P?BLICAS (GET/LET)
'******************************************************************************

Public Property Get idExpediente() As Long
    idExpediente = m_idExpediente
End Property
Public Property Let idExpediente(ByVal value As Long)
    m_idExpediente = value
End Property

Public Property Get Nemotecnico() As String
    Nemotecnico = m_Nemotecnico
End Property
Public Property Let Nemotecnico(ByVal value As String)
    m_Nemotecnico = value
End Property

Public Property Get Titulo() As String
    Titulo = m_Titulo
End Property
Public Property Let Titulo(ByVal value As String)
    m_Titulo = value
End Property

Public Property Get ResponsableCalidad() As String
    ResponsableCalidad = m_ResponsableCalidad
End Property
Public Property Let ResponsableCalidad(ByVal value As String)
    m_ResponsableCalidad = value
End Property

Public Property Get ResponsableTecnico() As String
    ResponsableTecnico = m_ResponsableTecnico
End Property
Public Property Let ResponsableTecnico(ByVal value As String)
    m_ResponsableTecnico = value
End Property

Public Property Get Pecal() As String
    Pecal = m_Pecal
End Property
Public Property Let Pecal(ByVal value As String)
    m_Pecal = value
End Property

Public Property Get NumeroExpediente() As String
    NumeroExpediente = m_NumeroExpediente
End Property
Public Property Let NumeroExpediente(ByVal value As String)
    m_NumeroExpediente = value
End Property

Public Property Get EstadoExpediente() As String
    EstadoExpediente = m_EstadoExpediente
End Property
Public Property Let EstadoExpediente(ByVal value As String)
    m_EstadoExpediente = value
End Property

Public Property Get EmailResponsable() As String
    EmailResponsable = m_EmailResponsable
End Property
Public Property Let EmailResponsable(ByVal value As String)
    m_EmailResponsable = value
End Property

Public Property Get JefeProyecto() As String
    JefeProyecto = m_JefeProyecto
End Property
Public Property Let JefeProyecto(ByVal value As String)
    m_JefeProyecto = value
End Property

Public Property Get ContratistaPrincipal() As String
    ContratistaPrincipal = m_ContratistaPrincipal
End Property
Public Property Let ContratistaPrincipal(ByVal value As String)
    m_ContratistaPrincipal = value
End Property

Public Property Get FechaInicio() As Date
    FechaInicio = m_FechaInicio
End Property
Public Property Let FechaInicio(ByVal value As Date)
    m_FechaInicio = value
End Property

Public Property Get FechaFinPrevista() As Date
    FechaFinPrevista = m_FechaFinPrevista
End Property
Public Property Let FechaFinPrevista(ByVal value As Date)
    m_FechaFinPrevista = value
End Property

Public Property Get FechaFinContrato() As Date
    FechaFinContrato = m_FechaFinContrato
End Property
Public Property Let FechaFinContrato(ByVal value As Date)
    m_FechaFinContrato = value
End Property

Public Property Get FechaFinGarantia() As Date
    FechaFinGarantia = m_FechaFinGarantia
End Property
Public Property Let FechaFinGarantia(ByVal value As Date)
    m_FechaFinGarantia = value
End Property

Public Property Get descripcion() As String
    descripcion = m_Descripcion
End Property
Public Property Let descripcion(ByVal value As String)
    m_Descripcion = value
End Property

Public Property Get fechaCreacion() As Date
    fechaCreacion = m_fechaCreacion
End Property
Public Property Let fechaCreacion(ByVal value As Date)
    m_fechaCreacion = value
End Property

Public Property Get Estado() As String
    Estado = m_Estado
End Property
Public Property Let Estado(ByVal value As String)
    m_Estado = value
End Property

Public Property Get IdUsuarioCreador() As Long
    IdUsuarioCreador = m_IdUsuarioCreador
End Property
Public Property Let IdUsuarioCreador(ByVal value As Long)
    m_IdUsuarioCreador = value
End Property

Public Property Get NombreUsuarioCreador() As String
    NombreUsuarioCreador = m_NombreUsuarioCreador
End Property
Public Property Let NombreUsuarioCreador(ByVal value As String)
    m_NombreUsuarioCreador = value
End Property

'******************************************************************************
' M?TODOS P?BLICOS
'******************************************************************************

' M?todo para validar completitud de datos
Public Function EsValido() As Boolean
    EsValido = (Len(NumeroExpediente) > 0 And _
                Len(ResponsableCalidad) > 0 And _
                Len(EmailResponsable) > 0)
End Function

