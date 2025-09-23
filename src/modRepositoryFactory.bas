Attribute VB_Name = "modRepositoryFactory"
Option Compare Database
Option Explicit


' =====================================================
' FACTORY: modRepositoryFactory
' DESCRIPCIÓN: Crea instancias de TODOS los repositorios del sistema.
' PATRÓN ARQUITECTÓNICO: Factory Testeable con Parámetros Opcionales.
' ESTÁNDAR: Oro - Versión Completa
' =====================================================

' Flag para alternar entre implementaciones reales y mocks
Public Const DEV_MODE As Boolean = True

' --- Métodos de Creación de Repositorios ---

Public Function CreateAuthRepository(Optional ByVal config As IConfig = Nothing) As IAuthRepository
    On Error GoTo ErrorHandler
    
    Dim effectiveConfig As IConfig
    If config Is Nothing Then
        ' Si no se pasa una configuración, usar la global por defecto
        Set effectiveConfig = modTestContext.GetTestConfig()
    Else
        ' Si se pasa una configuración (desde un test), usarla
        Set effectiveConfig = config
    End If
    
    Dim ErrorHandler As IErrorHandlerService
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService(effectiveConfig)
    
    Dim repoImpl As New CAuthRepository
    repoImpl.Initialize effectiveConfig, ErrorHandler
    Set CreateAuthRepository = repoImpl
    Exit Function
ErrorHandler:
    Debug.Print "Error crítico en modRepositoryFactory.CreateAuthRepository: " & Err.Description
End Function

Public Function CreateExpedienteRepository(Optional ByVal config As IConfig = Nothing) As IExpedienteRepository
    On Error GoTo ErrorHandler
    
    Dim effectiveConfig As IConfig
    If config Is Nothing Then
        ' Si no se pasa una configuración, usar la global por defecto
        Set effectiveConfig = modTestContext.GetTestConfig()
    Else
        ' Si se pasa una configuración (desde un test), usarla
        Set effectiveConfig = config
    End If
    
    Dim ErrorHandler As IErrorHandlerService
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService(effectiveConfig)
    
    Dim repoImpl As New CExpedienteRepository
    repoImpl.Initialize effectiveConfig, ErrorHandler
    Set CreateExpedienteRepository = repoImpl
    Exit Function
ErrorHandler:
    Debug.Print "Error crítico en modRepositoryFactory.CreateExpedienteRepository: " & Err.Description
End Function

Public Function CreateMapeoRepository(Optional ByVal config As IConfig = Nothing) As IMapeoRepository
    On Error GoTo ErrorHandler
    
    Dim effectiveConfig As IConfig
    If config Is Nothing Then
        ' Si no se pasa una configuración, usar la global por defecto
        Set effectiveConfig = modTestContext.GetTestConfig()
    Else
        ' Si se pasa una configuración (desde un test), usarla
        Set effectiveConfig = config
    End If
    
    Dim ErrorHandler As IErrorHandlerService
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService(effectiveConfig)
    
    Dim repoImpl As New CMapeoRepository
    repoImpl.Initialize effectiveConfig, ErrorHandler
    Set CreateMapeoRepository = repoImpl
    Exit Function
ErrorHandler:
    Debug.Print "Error crítico en modRepositoryFactory.CreateMapeoRepository: " & Err.Description
End Function

Public Function CreateNotificationRepository(Optional ByVal config As IConfig = Nothing) As INotificationRepository
    On Error GoTo ErrorHandler
    
    Dim effectiveConfig As IConfig
    If config Is Nothing Then
        ' Si no se pasa una configuración, usar la global por defecto
        Set effectiveConfig = modTestContext.GetTestConfig()
    Else
        ' Si se pasa una configuración (desde un test), usarla
        Set effectiveConfig = config
    End If
    
    Dim ErrorHandler As IErrorHandlerService
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService(effectiveConfig)
    
    Dim repoImpl As New CNotificationRepository
    repoImpl.Initialize effectiveConfig, ErrorHandler
    Set CreateNotificationRepository = repoImpl
    Exit Function
ErrorHandler:
    Debug.Print "Error crítico en modRepositoryFactory.CreateNotificationRepository: " & Err.Description
End Function

Public Function CreateOperationRepository(Optional ByVal config As IConfig = Nothing) As IOperationRepository
    On Error GoTo ErrorHandler
    
    Dim effectiveConfig As IConfig
    If config Is Nothing Then
        ' Si no se pasa una configuración, usar la global por defecto
        Set effectiveConfig = modTestContext.GetTestConfig()
    Else
        ' Si se pasa una configuración (desde un test), usarla
        Set effectiveConfig = config
    End If
    
    Dim ErrorHandler As IErrorHandlerService
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService(effectiveConfig)
    
    Dim repoImpl As New COperationRepository
    repoImpl.Initialize effectiveConfig, ErrorHandler
    Set CreateOperationRepository = repoImpl
    Exit Function
ErrorHandler:
    Debug.Print "Error crítico en modRepositoryFactory.CreateOperationRepository: " & Err.Description
End Function

Public Function CreateSolicitudRepository(Optional ByVal config As IConfig = Nothing) As ISolicitudRepository
    On Error GoTo ErrorHandler
    
    Dim effectiveConfig As IConfig
    If config Is Nothing Then
        ' Si no se pasa una configuración, usar la global por defecto
        Set effectiveConfig = modTestContext.GetTestConfig()
    Else
        ' Si se pasa una configuración (desde un test), usarla
        Set effectiveConfig = config
    End If
    
    Dim ErrorHandler As IErrorHandlerService
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService(effectiveConfig)
    
    Dim repoImpl As New CSolicitudRepository
    repoImpl.Initialize effectiveConfig, ErrorHandler
    Set CreateSolicitudRepository = repoImpl
    Exit Function
ErrorHandler:
    Debug.Print "Error crítico en modRepositoryFactory.CreateSolicitudRepository: " & Err.Description
End Function

Public Function CreateWorkflowRepository(Optional ByVal config As IConfig = Nothing) As IWorkflowRepository
    On Error GoTo ErrorHandler
    
    Dim effectiveConfig As IConfig
    If config Is Nothing Then
        ' Si no se pasa una configuración, usar la global por defecto
        Set effectiveConfig = modTestContext.GetTestConfig()
    Else
        ' Si se pasa una configuración (desde un test), usarla
        Set effectiveConfig = config
    End If
    
    Dim ErrorHandler As IErrorHandlerService
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService(effectiveConfig)
    
    Dim repoImpl As New CWorkflowRepository
    repoImpl.Initialize effectiveConfig, ErrorHandler
    Set CreateWorkflowRepository = repoImpl
    Exit Function
ErrorHandler:
    Debug.Print "Error crítico en modRepositoryFactory.CreateWorkflowRepository: " & Err.Description
End Function
