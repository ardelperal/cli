VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EEstado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private m_idEstado As Long
Private m_nombreEstado As String
Private m_Descripcion As String
Private m_esEstadoInicial As Boolean
Private m_esEstadoFinal As Boolean
Private m_orden As Integer

' Propiedades p?blicas
Public Property Get IdEstado() As Long
    IdEstado = m_idEstado
End Property

Public Property Let IdEstado(ByVal value As Long)
    m_idEstado = value
End Property

Public Property Get NombreEstado() As String
    NombreEstado = m_nombreEstado
End Property

Public Property Let NombreEstado(ByVal value As String)
    m_nombreEstado = value
End Property

Public Property Get descripcion() As String
    descripcion = m_Descripcion
End Property

Public Property Let descripcion(ByVal value As String)
    m_Descripcion = value
End Property

Public Property Get EsEstadoInicial() As Boolean
    EsEstadoInicial = m_esEstadoInicial
End Property

Public Property Let EsEstadoInicial(ByVal value As Boolean)
    m_esEstadoInicial = value
End Property

Public Property Get EsEstadoFinal() As Boolean
    EsEstadoFinal = m_esEstadoFinal
End Property

Public Property Let EsEstadoFinal(ByVal value As Boolean)
    m_esEstadoFinal = value
End Property

Public Property Get Orden() As Integer
    Orden = m_orden
End Property

Public Property Let Orden(ByVal value As Integer)
    m_orden = value
End Property
