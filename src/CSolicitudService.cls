VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSolicitudService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements ISolicitudService

' Dependencias inyectadas
Private m_solicitudRepository As ISolicitudRepository
Private m_OperationLogger As IOperationLogger
Private m_ErrorHandler As IErrorHandlerService
Private m_authService As IAuthService
Private m_workflowService As IWorkflowService
Private m_isInitialized As Boolean

' ============================================================================
' INICIALIZACI?N Y CONFIGURACI?N
' ============================================================================

' Inicializa el servicio con sus dependencias
' @param solicitudRepository: Repositorio para operaciones de persistencia
' @param operationLogger: Logger para registrar operaciones
' @param errorHandler: Manejador de errores para logging
' @param authService: Servicio de autenticaci?n para obtener contexto del usuario
' @param workflowService: Servicio de workflow para validar transiciones
Public Sub Initialize(ByVal solicitudRepository As ISolicitudRepository, ByVal operationLogger As IOperationLogger, ByVal ErrorHandler As IErrorHandlerService, ByVal authService As IAuthService, ByVal workflowService As IWorkflowService)
    If solicitudRepository Is Nothing Then
        Err.Raise 5, "CSolicitudService.Initialize", "solicitudRepository no puede ser Nothing"
    End If
    
    If operationLogger Is Nothing Then
        Err.Raise 5, "CSolicitudService.Initialize", "operationLogger no puede ser Nothing"
    End If
    
    If ErrorHandler Is Nothing Then
        Err.Raise 5, "CSolicitudService.Initialize", "errorHandler no puede ser Nothing"
    End If
    
    If authService Is Nothing Then
        Err.Raise 5, "CSolicitudService.Initialize", "authService no puede ser Nothing"
    End If
    
    If workflowService Is Nothing Then
        Err.Raise 5, "CSolicitudService.Initialize", "workflowService no puede ser Nothing"
    End If
    
    Set m_solicitudRepository = solicitudRepository
    Set m_OperationLogger = operationLogger
    Set m_ErrorHandler = ErrorHandler
    Set m_authService = authService
    Set m_workflowService = workflowService
    m_isInitialized = True
End Sub

' ============================================================================
' IMPLEMENTACI?N DE ISolicitudService
' ============================================================================

' Crea una nueva solicitud con estado inicial "Borrador"
Private Function ISolicitudService_CreateSolicitud(ByVal paraExpediente As EExpediente) As ESolicitud
    
    If Not m_isInitialized Then
        Err.Raise 5, "CSolicitudService.CreateSolicitud", "El servicio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    If paraExpediente Is Nothing Then
        Err.Raise 5, "CSolicitudService.CreateSolicitud", "El par?metro paraExpediente no puede ser Nothing."
    End If
    
    If paraExpediente.idExpediente <= 0 Then
        Err.Raise vbObjectError + 1001, "CSolicitudService.CreateSolicitud", "El ID del expediente proporcionado no es v?lido."
    End If
    

    
    On Error GoTo ErrorHandler
    
    ' Crear nueva solicitud con estado inicial "Borrador"
    Dim nuevaSolicitud As ESolicitud
    Set nuevaSolicitud = New ESolicitud
    
    With nuevaSolicitud
        .idSolicitud = 0 ' Ser? asignado por el repositorio
        .idExpediente = paraExpediente.idExpediente
        .tipoSolicitud = "PC" ' Tipo por defecto
        .subTipoSolicitud = ""
        .codigoSolicitud = GenerateCodigoSolicitud(CStr(paraExpediente.idExpediente), "PC")
        .idEstadoInterno = 1 ' ID del estado Borrador en tbEstados
        .fechaCreacion = Now()
        .usuarioCreacion = m_authService.GetCurrentUserEmail() ' Usuario actual del sistema
        .usuarioModificacion = ""
    End With
    
    ' Delegar al repositorio para guardar
    If m_solicitudRepository.SaveSolicitud(nuevaSolicitud) > 0 Then
        Dim logEntry As EOperationLog
        Set logEntry = New EOperationLog
        With logEntry
            .tipoOperacion = "CREATE"
            .idEntidadAfectada = nuevaSolicitud.idSolicitud
            .usuario = m_authService.GetCurrentUserEmail()
            .resultado = "Success"
            .descripcion = "Nueva solicitud creada: ID=" & nuevaSolicitud.idSolicitud & ", Expediente=" & paraExpediente.idExpediente & ", Tipo=PC"
        End With
        m_OperationLogger.LogOperation logEntry
        
        Set ISolicitudService_CreateSolicitud = nuevaSolicitud
    Else
        Dim logErrorEntry As EOperationLog
        Set logErrorEntry = New EOperationLog
        With logErrorEntry
            .tipoOperacion = "CREATE"
            .idEntidadAfectada = paraExpediente.idExpediente
            .usuario = m_authService.GetCurrentUserEmail()
            .resultado = "Error"
            .descripcion = "Fallo al intentar crear la solicitud para el expediente ID=" & paraExpediente.idExpediente
        End With
        m_OperationLogger.LogOperation logErrorEntry
        Err.Raise 5, "CSolicitudService.CreateSolicitud", "No se pudo crear la solicitud"
    End If
    
    Exit Function
    
ErrorHandler:
    Dim logCrashEntry As EOperationLog
    Set logCrashEntry = New EOperationLog
    With logCrashEntry
        .tipoOperacion = "CREATE"
        .idEntidadAfectada = 0
        .usuario = m_authService.GetCurrentUserEmail()
        .resultado = "Crash"
        .detalles = "Error: " & Err.Description
    End With
    m_OperationLogger.LogOperation logCrashEntry
    m_ErrorHandler.LogError Err.Number, Err.Description, "CSolicitudService.CreateSolicitud"
    Err.Raise Err.Number, "CSolicitudService.CreateSolicitud", Err.Description
End Function

' Guarda una solicitud existente
Private Function ISolicitudService_SaveSolicitud(ByVal solicitud As ESolicitud) As Long
    
    If Not m_isInitialized Then
        Err.Raise 5, "CSolicitudService.SaveSolicitud", "El servicio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    If solicitud Is Nothing Then
        Err.Raise 5, "CSolicitudService.SaveSolicitud", "El par?metro solicitud no puede ser Nothing."
    End If
    
    If solicitud.idSolicitud <= 0 Then
        Err.Raise vbObjectError + 1002, "CSolicitudService.SaveSolicitud", "El m?todo SaveSolicitud solo puede usarse en solicitudes con un ID existente."
    End If
    
    On Error GoTo ErrorHandler
    
    ' Actualizar campos de modificaci?n
    solicitud.fechaModificacion = Now()
    solicitud.usuarioModificacion = m_authService.GetCurrentUserEmail()
    
    ' Delegar al repositorio para guardar
    Dim resultado As Long
    resultado = m_solicitudRepository.SaveSolicitud(solicitud)
    If resultado > 0 Then
        Dim logEntry As EOperationLog
        Set logEntry = New EOperationLog
        With logEntry
            .tipoOperacion = "SAVE"
            .idEntidadAfectada = solicitud.idSolicitud
            .usuario = m_authService.GetCurrentUserEmail()
            .resultado = "Success"
            .descripcion = "Solicitud guardada correctamente: ID=" & solicitud.idSolicitud
        End With
        m_OperationLogger.LogOperation logEntry
        ISolicitudService_SaveSolicitud = resultado
    Else
        Dim logErrorEntry As EOperationLog
        Set logErrorEntry = New EOperationLog
        With logErrorEntry
            .tipoOperacion = "SAVE"
            .idEntidadAfectada = solicitud.idSolicitud
            .usuario = m_authService.GetCurrentUserEmail()
            .resultado = "Error"
            .descripcion = "Error: No se pudo guardar la solicitud ID=" & solicitud.idSolicitud
        End With
        m_OperationLogger.LogOperation logErrorEntry
        ISolicitudService_SaveSolicitud = 0
    End If
    
    Exit Function
    
ErrorHandler:
    Dim logCrashEntry As EOperationLog
    Set logCrashEntry = New EOperationLog
    With logCrashEntry
        .tipoOperacion = "SAVE"
        .idEntidadAfectada = 0
        .usuario = m_authService.GetCurrentUserEmail()
        .resultado = "Crash"
        .detalles = "Error: " & Err.Description
    End With
    m_OperationLogger.LogOperation logCrashEntry
    m_ErrorHandler.LogError Err.Number, Err.Description, "CSolicitudService.SaveSolicitud"
    ISolicitudService_SaveSolicitud = 0
End Function

' ============================================================================
' M?TODOS PRIVADOS DE APOYO
' ============================================================================

' Genera un c?digo ?nico para la solicitud
' @param idExpediente: ID del expediente
' @param tipo: Tipo de solicitud
' @return: C?digo generado
Private Function GenerateCodigoSolicitud(ByVal idExpediente As String, ByVal tipo As String) As String
    Dim timestamp As String
    timestamp = Format(Now(), "yyyymmddhhnnss")
    
    GenerateCodigoSolicitud = tipo & "-" & idExpediente & "-" & timestamp
End Function

'==============================================================================
' M?TODOS P?BLICOS DE CONVENIENCIA
'==============================================================================

' M?todo p?blico de conveniencia que delega a la implementaci?n de la interfaz
Public Function CreateSolicitud(ByVal paraExpediente As EExpediente) As ESolicitud
    Set CreateSolicitud = ISolicitudService_CreateSolicitud(paraExpediente)
End Function

' M?todo p?blico de conveniencia que delega a la implementaci?n de la interfaz
Public Function SaveSolicitud(ByVal solicitud As ESolicitud) As Long
    SaveSolicitud = ISolicitudService_SaveSolicitud(solicitud)
End Function

' Obtiene una solicitud por su ID
Private Function ISolicitudService_ObtenerSolicitudPorId(ByVal SolicitudID As Long) As ESolicitud
    
    If Not m_isInitialized Then
        Err.Raise 5, "CSolicitudService.ObtenerSolicitudPorId", "El servicio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    If SolicitudID <= 0 Then
        Err.Raise 5, "CSolicitudService.ObtenerSolicitudPorId", "SolicitudID debe ser mayor que 0"
    End If
    
    On Error GoTo ErrorHandler
    
    ' Delegar al repositorio para obtener la solicitud
    Dim solicitud As ESolicitud
    Set solicitud = m_solicitudRepository.ObtenerSolicitudPorId(SolicitudID)
    
    If solicitud Is Nothing Then
        Dim logNotFoundEntry As EOperationLog
        Set logNotFoundEntry = New EOperationLog
        With logNotFoundEntry
            .tipoOperacion = "GET"
            .idEntidadAfectada = SolicitudID
            .usuario = m_authService.GetCurrentUserEmail()
            .resultado = "NotFound"
            .descripcion = "Solicitud no encontrada: ID=" & SolicitudID
        End With
        m_OperationLogger.LogOperation logNotFoundEntry
    Else
        Dim logSuccessEntry As EOperationLog
        Set logSuccessEntry = New EOperationLog
        With logSuccessEntry
            .tipoOperacion = "GET"
            .idEntidadAfectada = SolicitudID
            .usuario = m_authService.GetCurrentUserEmail()
            .resultado = "Success"
            .descripcion = "Solicitud obtenida correctamente: ID=" & SolicitudID
        End With
        m_OperationLogger.LogOperation logSuccessEntry
    End If
    
    Set ISolicitudService_ObtenerSolicitudPorId = solicitud
    Exit Function
    
ErrorHandler:
    Dim logCrashEntry As EOperationLog
    Set logCrashEntry = New EOperationLog
    With logCrashEntry
        .tipoOperacion = "GET"
        .idEntidadAfectada = SolicitudID
        .usuario = m_authService.GetCurrentUserEmail()
        .resultado = "Crash"
        .detalles = "Error: " & Err.Description
    End With
    m_OperationLogger.LogOperation logCrashEntry
    m_ErrorHandler.LogError Err.Number, Err.Description, "CSolicitudService.ObtenerSolicitudPorId"
    Err.Raise Err.Number, "CSolicitudService.ObtenerSolicitudPorId", Err.Description
End Function

' M?todo p?blico de conveniencia que delega a la implementaci?n de la interfaz
Public Function ObtenerSolicitudPorId(ByVal SolicitudID As Long) As ESolicitud
    Set ObtenerSolicitudPorId = ISolicitudService_ObtenerSolicitudPorId(SolicitudID)
End Function

' Cambia el estado de una solicitud validando la transici?n
Private Function ISolicitudService_CambiarEstadoSolicitud(ByVal SolicitudID As Long, ByVal nuevoEstadoId As Long) As Boolean
    
    If Not m_isInitialized Then
        Err.Raise 5, "CSolicitudService.CambiarEstadoSolicitud", "El servicio no ha sido inicializado. Llame a Initialize() primero."
    End If
    
    If SolicitudID <= 0 Then
        Err.Raise vbObjectError + 1003, "CSolicitudService.CambiarEstadoSolicitud", "El ID de la solicitud no es v?lido."
    End If

    If nuevoEstadoId <= 0 Then
        Err.Raise vbObjectError + 1004, "CSolicitudService.CambiarEstadoSolicitud", "El ID del nuevo estado no es v?lido."
    End If
    
    On Error GoTo ErrorHandler
    
    ' Obtener la solicitud actual
    Dim solicitud As ESolicitud
    Set solicitud = m_solicitudRepository.ObtenerSolicitudPorId(SolicitudID)
    
    If solicitud Is Nothing Then
        Dim logNotFoundEntry As EOperationLog
        Set logNotFoundEntry = New EOperationLog
        With logNotFoundEntry
            .tipoOperacion = "CHANGE_STATE"
            .idEntidadAfectada = SolicitudID
            .usuario = m_authService.GetCurrentUserEmail()
            .resultado = "Error"
            .descripcion = "Solicitud no encontrada: ID=" & SolicitudID
        End With
        m_OperationLogger.LogOperation logNotFoundEntry
        ISolicitudService_CambiarEstadoSolicitud = False
        Exit Function
    End If
    
    ' Obtener el estado actual y nuevo estado como strings
    Dim estadoActual As String
    Dim estadoDestino As String
    
    ' Mapear IDs a nombres de estados
    Select Case solicitud.idEstadoInterno
        Case 1: estadoActual = "Registrado"
        Case 2: estadoActual = "Desarrollo"
        Case 3: estadoActual = "Modificaci?n"
        Case 4: estadoActual = "Validaci?n"
        Case 5: estadoActual = "Revisi?n"
        Case 6: estadoActual = "Formalizaci?n"
        Case 7: estadoActual = "Aprobada"
        Case Else: estadoActual = "Desconocido"
    End Select
    
    Select Case nuevoEstadoId
        Case 1: estadoDestino = "Registrado"
        Case 2: estadoDestino = "Desarrollo"
        Case 3: estadoDestino = "Modificaci?n"
        Case 4: estadoDestino = "Validaci?n"
        Case 5: estadoDestino = "Revisi?n"
        Case 6: estadoDestino = "Formalizaci?n"
        Case 7: estadoDestino = "Aprobada"
        Case Else: estadoDestino = "Desconocido"
    End Select
    
    ' Obtener el rol del usuario actual (simplificado)
    Dim usuarioRol As String
    usuarioRol = "Calidad" ' Por defecto, en una implementaci?n real se obtendr?a del contexto del usuario
    
    ' Validar la transici?n usando el WorkflowService
    Dim esTransicionValida As Boolean
    esTransicionValida = m_workflowService.ValidateTransition(SolicitudID, estadoActual, estadoDestino, solicitud.tipoSolicitud, usuarioRol)
    
    If Not esTransicionValida Then
        Dim logTransitionEntry As EOperationLog
        Set logTransitionEntry = New EOperationLog
        With logTransitionEntry
            .tipoOperacion = "CHANGE_STATE"
            .idEntidadAfectada = SolicitudID
            .usuario = m_authService.GetCurrentUserEmail()
            .resultado = "Error"
            .descripcion = "Transici?n no permitida: " & estadoActual & " -> " & estadoDestino & " para rol " & usuarioRol
        End With
        m_OperationLogger.LogOperation logTransitionEntry
        m_ErrorHandler.LogError 5, "Transici?n de estado no permitida", "CSolicitudService.CambiarEstadoSolicitud"
        Err.Raise 5, "CSolicitudService.CambiarEstadoSolicitud", "Transici?n de estado no permitida"
    End If
    
    ' Si la transici?n es v?lida, actualizar el estado
    solicitud.idEstadoInterno = nuevoEstadoId
    solicitud.fechaModificacion = Now()
    solicitud.usuarioModificacion = m_authService.GetCurrentUserEmail()
    
    ' Guardar la solicitud actualizada
    Dim resultado As Long
    resultado = m_solicitudRepository.SaveSolicitud(solicitud)
    
    If resultado > 0 Then
        Dim logSuccessEntry As EOperationLog
        Set logSuccessEntry = New EOperationLog
        With logSuccessEntry
            .tipoOperacion = "CHANGE_STATE"
            .idEntidadAfectada = SolicitudID
            .usuario = m_authService.GetCurrentUserEmail()
            .resultado = "Success"
            .descripcion = "Estado cambiado exitosamente: " & estadoActual & " -> " & estadoDestino
        End With
        m_OperationLogger.LogOperation logSuccessEntry
        ISolicitudService_CambiarEstadoSolicitud = True
    Else
        Dim logSaveErrorEntry As EOperationLog
        Set logSaveErrorEntry = New EOperationLog
        With logSaveErrorEntry
            .tipoOperacion = "CHANGE_STATE"
            .idEntidadAfectada = SolicitudID
            .usuario = m_authService.GetCurrentUserEmail()
            .resultado = "Error"
            .descripcion = "Error al guardar solicitud con nuevo estado"
        End With
        m_OperationLogger.LogOperation logSaveErrorEntry
        ISolicitudService_CambiarEstadoSolicitud = False
    End If
    
    Exit Function
    
ErrorHandler:
    Dim logCrashEntry As EOperationLog
    Set logCrashEntry = New EOperationLog
    With logCrashEntry
        .tipoOperacion = "CHANGE_STATE"
        .idEntidadAfectada = SolicitudID
        .usuario = m_authService.GetCurrentUserEmail()
        .resultado = "Crash"
        .detalles = "Error: " & Err.Description
    End With
    m_OperationLogger.LogOperation logCrashEntry
    m_ErrorHandler.LogError Err.Number, Err.Description, "CSolicitudService.CambiarEstadoSolicitud"
    ISolicitudService_CambiarEstadoSolicitud = False
End Function

' M?todo p?blico de conveniencia que delega a la implementaci?n de la interfaz
Public Function CambiarEstadoSolicitud(ByVal SolicitudID As Long, ByVal nuevoEstadoId As Long) As Boolean
    CambiarEstadoSolicitud = ISolicitudService_CambiarEstadoSolicitud(SolicitudID, nuevoEstadoId)
End Function
