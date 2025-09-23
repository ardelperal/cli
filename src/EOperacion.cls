VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EOperacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private m_idOperacion As Long
Private m_FechaHora As Date
Private m_Usuario As String
Private m_TipoOperacion As String
Private m_Entidad As String
Private m_idEntidad As Long
Private m_Descripcion As String
Private m_Resultado As String
Private m_Detalles As String

' Propiedades p?blicas
Public Property Get IdOperacion() As Long
    IdOperacion = m_idOperacion
End Property

Public Property Let IdOperacion(ByVal value As Long)
    m_idOperacion = value
End Property

Public Property Get fechaHora() As Date
    fechaHora = m_FechaHora
End Property

Public Property Let fechaHora(ByVal value As Date)
    m_FechaHora = value
End Property

Public Property Get usuario() As String
    usuario = m_Usuario
End Property

Public Property Let usuario(ByVal value As String)
    m_Usuario = value
End Property

Public Property Get tipoOperacion() As String
    tipoOperacion = m_TipoOperacion
End Property

Public Property Let tipoOperacion(ByVal value As String)
    m_TipoOperacion = value
End Property

Public Property Get entidad() As String
    entidad = m_Entidad
End Property

Public Property Let entidad(ByVal value As String)
    m_Entidad = value
End Property

Public Property Get IdEntidad() As Long
    IdEntidad = m_idEntidad
End Property

Public Property Let IdEntidad(ByVal value As Long)
    m_idEntidad = value
End Property

Public Property Get descripcion() As String
    descripcion = m_Descripcion
End Property

Public Property Let descripcion(ByVal value As String)
    m_Descripcion = value
End Property

Public Property Get resultado() As String
    resultado = m_Resultado
End Property

Public Property Let resultado(ByVal value As String)
    m_Resultado = value
End Property

Public Property Get detalles() As String
    detalles = m_Detalles
End Property

Public Property Let detalles(ByVal value As String)
    m_Detalles = value
End Property
