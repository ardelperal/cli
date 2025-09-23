VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAuthService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IAuthService

Private configSvc As IConfig
Private m_OperationLogger As IOperationLogger
Private m_AuthRepository As IAuthRepository
Private m_ErrorHandler As IErrorHandlerService

Private Sub Class_Initialize()
    ' Inicializaci?n sin dependencias
End Sub

Public Sub Initialize(config As IConfig, operationLogger As IOperationLogger, authRepository As IAuthRepository, ErrorHandler As IErrorHandlerService)
    On Error GoTo ErrorHandler
    
    Set configSvc = config
    Set m_OperationLogger = operationLogger
    Set m_AuthRepository = authRepository
    Set m_ErrorHandler = ErrorHandler
    
    Exit Sub
    
ErrorHandler:
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CAuthService.Initialize"
    End If
End Sub

Private Function IAuthService_GetUserRole(ByVal userEmail As String) As UserRole
    On Error GoTo ErrorHandler

    ' Logging de auditor?a: Consulta de rol de usuario
    If Not m_OperationLogger Is Nothing Then
        Dim logEntry As EOperationLog
        Set logEntry = New EOperationLog
        With logEntry
            .tipoOperacion = "ROLE_QUERY"
            .idEntidadAfectada = 0
            .usuario = userEmail
            .resultado = "Info"
            .descripcion = "Iniciando consulta de rol para usuario: " & userEmail
        End With
        m_OperationLogger.LogOperation logEntry
    End If

    Dim role As UserRole
    Dim authData As EAuthData

    role = RolDesconocido

    ' Obtener todos los datos de autenticaci?n en una sola llamada
    Set authData = m_AuthRepository.GetUserAuthData(userEmail)

    If Not authData Is Nothing Then
        ' Verificar si el usuario existe
        If authData.UserExists Then
            ' Determinar el rol bas?ndose en las propiedades de EAuthData
            ' Prioridad: Administrador Global > Administrador App > Calidad > T?cnico
            If authData.IsGlobalAdmin Then
                role = RolAdmin
            ElseIf authData.IsAppAdmin Then
                role = RolAdmin
            ElseIf authData.IsCalidad Then
                role = RolCalidad
            ElseIf authData.IsTecnico Then
                role = RolTecnico
            Else
                role = RolDesconocido
            End If
        Else
            ' Usuario no encontrado
            role = RolDesconocido
        End If
    End If
    
    Set authData = Nothing

    ' Logging de auditor?a: Resultado de consulta de rol
    If Not m_OperationLogger Is Nothing Then
        Dim roleText As String
        Select Case role
            Case RolAdmin: roleText = "Administrador"
            Case RolCalidad: roleText = "Calidad"
            Case RolTecnico: roleText = "T?cnico"
            Case Else: roleText = "Desconocido"
        End Select
        Dim logResultEntry As EOperationLog
        Set logResultEntry = New EOperationLog
        With logResultEntry
            .tipoOperacion = "ROLE_QUERY"
            .idEntidadAfectada = 0
            .usuario = userEmail
            .resultado = "Success"
            .descripcion = "Rol asignado a " & userEmail & ": " & roleText
        End With
        m_OperationLogger.LogOperation logResultEntry
    End If
    
    IAuthService_GetUserRole = role
    Exit Function

ErrorHandler:
    IAuthService_GetUserRole = RolDesconocido
    If Not authData Is Nothing Then
        Set authData = Nothing
    End If
    If Not m_OperationLogger Is Nothing Then
        Dim logErrorEntry As EOperationLog
        Set logErrorEntry = New EOperationLog
        With logErrorEntry
            .tipoOperacion = "ROLE_QUERY"
            .idEntidadAfectada = 0
            .usuario = userEmail
            .resultado = "Crash"
            .descripcion = "Error fatal consultando rol para " & userEmail
            .detalles = Err.Description
        End With
        m_OperationLogger.LogOperation logErrorEntry
    End If
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CAuthService.GetUserRole"
    End If
End Function

Private Function IAuthService_AuthenticateUser(ByVal userEmail As String, ByVal password As String) As Boolean
    ' TODO: Implementar la l?gica real de autenticaci?n.
    ' Por ahora, devolvemos False para que compile.
    IAuthService_AuthenticateUser = False
End Function

' M?todo p?blico para acceso directo
Public Function GetUserRole(ByVal userEmail As String) As UserRole
    GetUserRole = IAuthService_GetUserRole(userEmail)
End Function

Public Function AuthenticateUser(ByVal userEmail As String, ByVal password As String) As Boolean
    AuthenticateUser = IAuthService_AuthenticateUser(userEmail, password)
End Function

' Implementaci?n del nuevo m?todo de la interfaz
Private Function IAuthService_GetCurrentUserEmail() As String
    IAuthService_GetCurrentUserEmail = configSvc.GetUsuarioActual()
End Function

' M?todo p?blico de conveniencia
Public Function GetCurrentUserEmail() As String
    GetCurrentUserEmail = IAuthService_GetCurrentUserEmail()
End Function
