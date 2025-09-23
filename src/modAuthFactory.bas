Attribute VB_Name = "modAuthFactory"
Option Compare Database
Option Explicit


Public Function CreateAuthService(Optional ByVal config As IConfig = Nothing) As IAuthService
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
    
    Dim operationLogger As IOperationLogger
    Set operationLogger = modOperationLoggerFactory.CreateOperationLogger(effectiveConfig)
    
    Dim authRepo As IAuthRepository
    Set authRepo = modRepositoryFactory.CreateAuthRepository(effectiveConfig) ' CORRECTO
    
    Dim authSvc As New CAuthService
    authSvc.Initialize effectiveConfig, operationLogger, authRepo, ErrorHandler
    
    Set CreateAuthService = authSvc
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error en modAuthFactory.CreateAuthService: " & Err.Description
    Set CreateAuthService = Nothing
End Function
