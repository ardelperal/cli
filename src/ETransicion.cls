VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ETransicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private m_idTransicion As Long
Private m_idEstadoOrigen As Long
Private m_idEstadoDestino As Long
Private m_rolRequerido As String
Private m_condiciones As String
Private m_accionesPost As String
Private m_activa As Boolean

' Propiedades p?blicas
Public Property Get IdTransicion() As Long
    IdTransicion = m_idTransicion
End Property

Public Property Let IdTransicion(ByVal value As Long)
    m_idTransicion = value
End Property

Public Property Get IdEstadoOrigen() As Long
    IdEstadoOrigen = m_idEstadoOrigen
End Property

Public Property Let IdEstadoOrigen(ByVal value As Long)
    m_idEstadoOrigen = value
End Property

Public Property Get IdEstadoDestino() As Long
    IdEstadoDestino = m_idEstadoDestino
End Property

Public Property Let IdEstadoDestino(ByVal value As Long)
    m_idEstadoDestino = value
End Property

Public Property Get RolRequerido() As String
    RolRequerido = m_rolRequerido
End Property

Public Property Let RolRequerido(ByVal value As String)
    m_rolRequerido = value
End Property

Public Property Get Condiciones() As String
    Condiciones = m_condiciones
End Property

Public Property Let Condiciones(ByVal value As String)
    m_condiciones = value
End Property

Public Property Get AccionesPost() As String
    AccionesPost = m_accionesPost
End Property

Public Property Let AccionesPost(ByVal value As String)
    m_accionesPost = value
End Property

Public Property Get Activa() As Boolean
    Activa = m_activa
End Property

Public Property Let Activa(ByVal value As Boolean)
    m_activa = value
End Property
