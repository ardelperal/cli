VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAppManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IAppManager

' =====================================================
' CLASE: CAppManager
' PROPOSITO: Implementaci?n de la gesti?n centralizada de la aplicaci?n
' AUTOR: CONDOR-Expert
' FECHA: 2025-01-14
' =====================================================

' Variables privadas para dependencias
Private m_authService As IAuthService
Private m_Config As IConfig
Private m_ErrorHandler As IErrorHandlerService
Private m_isInitialized As Boolean
Private m_CurrentUserRole As UserRole

' Definir constante de compilacion condicional para modo desarrollo


' =====================================================
' METODO: Initialize (IAppManager)
' =====================================================
Private Sub IAppManager_Initialize(ByVal authService As IAuthService, ByVal config As IConfig, ByVal ErrorHandler As IErrorHandlerService)
    Set m_authService = authService
    Set m_Config = config
    Set m_ErrorHandler = ErrorHandler
    m_isInitialized = False
    m_CurrentUserRole = RolDesconocido
End Sub

' =====================================================
' METODO: StartApplication (IAppManager)
' =====================================================
Private Function IAppManager_StartApplication(Optional ByVal userEmail As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    ' 1. Obtener el email del usuario actual
    Dim finalUserEmail As String
    If userEmail <> "" Then
        finalUserEmail = userEmail ' Usar el email proporcionado para testing
    Else
        finalUserEmail = GetCurrentUserEmail() ' Usar la l?gica normal de producci?n
    End If

    If finalUserEmail = "" Then
        m_ErrorHandler.LogError vbObjectError + 514, "No se pudo obtener el email del usuario para la autenticaci?n.", "CAppManager.StartApplication", "CRITICAL_STARTUP_FAILURE"
        IAppManager_StartApplication = False
        Exit Function
    End If

    ' 2. Determinar el rol del usuario y guardarlo
    m_CurrentUserRole = m_authService.GetUserRole(finalUserEmail)

    ' 3. Verificar que la configuraci?n se haya cargado correctamente
    If Not m_Config.GetValue("IsInitialized") Then
        ' El error espec?fico ya ha sido logueado por CConfig.ValidateConfiguration
        IAppManager_StartApplication = False
        Exit Function
    End If

    ' 4. Continuar con la inicializaci?n seg?n el rol
    Select Case m_CurrentUserRole
        Case RolAdmin
            Debug.Print "Usuario autenticado como Administrador."
            ' TODO: Implementar inicializaci?n espec?fica para Admin
        Case RolCalidad
            Debug.Print "Usuario autenticado como Calidad."
            ' TODO: Implementar inicializaci?n espec?fica para Calidad
        Case RolTecnico
            Debug.Print "Usuario autenticado como T?cnico."
            ' TODO: Implementar inicializaci?n espec?fica para T?cnico
        Case RolDesconocido
            m_ErrorHandler.LogError vbObjectError + 515, "El usuario '" & finalUserEmail & "' no tiene un rol definido en el sistema.", "CAppManager.StartApplication", "CRITICAL_AUTH_FAILURE"
            IAppManager_StartApplication = False
            Exit Function
    End Select
    
    m_isInitialized = True
    IAppManager_StartApplication = True
    Exit Function
    
ErrorHandler:
    m_ErrorHandler.LogError Err.Number, Err.Description, "CAppManager.StartApplication", "CRITICAL_STARTUP_ERROR"
    IAppManager_StartApplication = False
End Function

' =====================================================
' METODO: GetCurrentUserRole (IAppManager)
' =====================================================
Private Function IAppManager_GetCurrentUserRole() As UserRole
    IAppManager_GetCurrentUserRole = m_CurrentUserRole
End Function

' =====================================================
' METODO: IsApplicationInitialized (IAppManager)
' =====================================================
Private Function IAppManager_IsApplicationInitialized() As Boolean
    IAppManager_IsApplicationInitialized = m_isInitialized
End Function

' =====================================================
' METODO: Ping (IAppManager)
' =====================================================
Private Function IAppManager_Ping() As String
    IAppManager_Ping = "Pong"
End Function

' =====================================================
' METODOS PUBLICOS DE CONVENIENCIA (Lecci?n 24)
' =====================================================
Public Sub Initialize(ByVal authService As IAuthService, ByVal config As IConfig, ByVal ErrorHandler As IErrorHandlerService)
    IAppManager_Initialize authService, config, ErrorHandler
End Sub

Public Function StartApplication(Optional ByVal userEmail As String = "") As Boolean
    StartApplication = IAppManager_StartApplication(userEmail)
End Function

Public Function GetCurrentUserRole() As UserRole
    GetCurrentUserRole = IAppManager_GetCurrentUserRole()
End Function

Public Function IsApplicationInitialized() As Boolean
    IsApplicationInitialized = IAppManager_IsApplicationInitialized()
End Function

Public Function Ping() As String
    Ping = IAppManager_Ping()
End Function

' =====================================================
' FUNCION PRIVADA: GetCurrentUserEmail
' PROPOSITO: Obtener el email del usuario actual con capacidad de suplantaci?n en desarrollo
' =====================================================
Private Function GetCurrentUserEmail() As String
    On Error GoTo ErrorHandler
    
    ' Usar el usuario actual de Windows
    GetCurrentUserEmail = Environ("USERNAME") & "@" & m_Config.GetValue("DEFAULT_DOMAIN")
    
    Exit Function
    
ErrorHandler:
    m_ErrorHandler.LogError Err.Number, Err.Description, "CAppManager.GetCurrentUserEmail"
    GetCurrentUserEmail = "unknown@domain.com"
End Function

'----------------------------------
' PRUEBA
'-------------------------------------

