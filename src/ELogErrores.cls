VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ELogErrores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private m_idLogError As Long
Private m_FechaHora As Date
Private m_Usuario As String
Private m_modulo As String
Private m_procedimiento As String
Private m_numeroError As Long
Private m_descripcionError As String
Private m_contexto As String

' Propiedades p?blicas
Public Property Get IdLogError() As Long
    IdLogError = m_idLogError
End Property

Public Property Let IdLogError(ByVal value As Long)
    m_idLogError = value
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

Public Property Get Modulo() As String
    Modulo = m_modulo
End Property

Public Property Let Modulo(ByVal value As String)
    m_modulo = value
End Property

Public Property Get Procedimiento() As String
    Procedimiento = m_procedimiento
End Property

Public Property Let Procedimiento(ByVal value As String)
    m_procedimiento = value
End Property

Public Property Get NumeroError() As Long
    NumeroError = m_numeroError
End Property

Public Property Let NumeroError(ByVal value As Long)
    m_numeroError = value
End Property

Public Property Get DescripcionError() As String
    DescripcionError = m_descripcionError
End Property

Public Property Let DescripcionError(ByVal value As String)
    m_descripcionError = value
End Property

Public Property Get Contexto() As String
    Contexto = m_contexto
End Property

Public Property Let Contexto(ByVal value As String)
    m_contexto = value
End Property
