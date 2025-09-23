VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EAuthData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private m_UserExists As Boolean
Private m_IsGlobalAdmin As Boolean
Private m_IsAppAdmin As Boolean
Private m_IsCalidad As Boolean
Private m_IsTecnico As Boolean

' Constructor por defecto
Private Sub Class_Initialize()
    m_UserExists = False
    m_IsGlobalAdmin = False
    m_IsAppAdmin = False
    m_IsCalidad = False
    m_IsTecnico = False
End Sub

' Propiedades p?blicas
Public Property Get UserExists() As Boolean
    UserExists = m_UserExists
End Property

Public Property Let UserExists(ByVal value As Boolean)
    m_UserExists = value
End Property

Public Property Get IsGlobalAdmin() As Boolean
    IsGlobalAdmin = m_IsGlobalAdmin
End Property

Public Property Let IsGlobalAdmin(ByVal value As Boolean)
    m_IsGlobalAdmin = value
End Property

Public Property Get IsAppAdmin() As Boolean
    IsAppAdmin = m_IsAppAdmin
End Property

Public Property Let IsAppAdmin(ByVal value As Boolean)
    m_IsAppAdmin = value
End Property

Public Property Get IsCalidad() As Boolean
    IsCalidad = m_IsCalidad
End Property

Public Property Let IsCalidad(ByVal value As Boolean)
    m_IsCalidad = value
End Property

Public Property Get IsTecnico() As Boolean
    IsTecnico = m_IsTecnico
End Property

Public Property Let IsTecnico(ByVal value As Boolean)
    m_IsTecnico = value
End Property
