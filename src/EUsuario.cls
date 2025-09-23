VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private m_ID As Long
Private m_Email As String
Private m_NombreCompleto As String
Private m_Rol As UserRole
Private m_EsAdministrador As Boolean
Private m_EsUsuarioCalidad As Boolean
Private m_EsUsuarioTecnico As Boolean
Private m_VersionInstalada As String
Private m_UltimoAcceso As Date
Private m_FechaRegistro As Date
Private m_Permisos As String

'******************************************************************************
' PROPIEDADES P?BLICAS (GET/LET)
'******************************************************************************

Public Property Get ID() As Long
    ID = m_ID
End Property
Public Property Let ID(ByVal value As Long)
    m_ID = value
End Property

Public Property Get Email() As String
    Email = m_Email
End Property
Public Property Let Email(ByVal value As String)
    m_Email = value
End Property

Public Property Get NombreCompleto() As String
    NombreCompleto = m_NombreCompleto
End Property
Public Property Let NombreCompleto(ByVal value As String)
    m_NombreCompleto = value
End Property

Public Property Get Rol() As UserRole
    Rol = m_Rol
End Property
Public Property Let Rol(ByVal value As UserRole)
    m_Rol = value
End Property

Public Property Get EsAdministrador() As Boolean
    EsAdministrador = m_EsAdministrador
End Property
Public Property Let EsAdministrador(ByVal value As Boolean)
    m_EsAdministrador = value
End Property

Public Property Get EsUsuarioCalidad() As Boolean
    EsUsuarioCalidad = m_EsUsuarioCalidad
End Property
Public Property Let EsUsuarioCalidad(ByVal value As Boolean)
    m_EsUsuarioCalidad = value
End Property

Public Property Get EsUsuarioTecnico() As Boolean
    EsUsuarioTecnico = m_EsUsuarioTecnico
End Property
Public Property Let EsUsuarioTecnico(ByVal value As Boolean)
    m_EsUsuarioTecnico = value
End Property

Public Property Get VersionInstalada() As String
    VersionInstalada = m_VersionInstalada
End Property
Public Property Let VersionInstalada(ByVal value As String)
    m_VersionInstalada = value
End Property

Public Property Get UltimoAcceso() As Date
    UltimoAcceso = m_UltimoAcceso
End Property
Public Property Let UltimoAcceso(ByVal value As Date)
    m_UltimoAcceso = value
End Property

Public Property Get FechaRegistro() As Date
    FechaRegistro = m_FechaRegistro
End Property
Public Property Let FechaRegistro(ByVal value As Date)
    m_FechaRegistro = value
End Property

Public Property Get Permisos() As String
    Permisos = m_Permisos
End Property
Public Property Let Permisos(ByVal value As String)
    m_Permisos = value
End Property

'******************************************************************************
' M?TODOS P?BLICOS
'******************************************************************************

' Valida si el usuario tiene datos m?nimos requeridos
Public Function EsValido() As Boolean
    EsValido = (Len(Email) > 0 And _
                InStr(Email, "@") > 1 And _
                Rol <> RolDesconocido)
End Function

' Determina si el usuario tiene permisos de administrador
Public Function TienePermisosAdmin() As Boolean
    TienePermisosAdmin = (EsAdministrador Or Rol = RolAdmin)
End Function

' Determina si el usuario puede crear solicitudes
Public Function PuedeCrearSolicitudes() As Boolean
    PuedeCrearSolicitudes = (EsUsuarioCalidad Or Rol = RolCalidad Or TienePermisosAdmin())
End Function

' Determina si el usuario puede consultar solicitudes
Public Function PuedeConsultarSolicitudes() As Boolean
    PuedeConsultarSolicitudes = (EsUsuarioTecnico Or Rol = RolTecnico Or _
                                EsUsuarioCalidad Or Rol = RolCalidad Or _
                                TienePermisosAdmin())
End Function

