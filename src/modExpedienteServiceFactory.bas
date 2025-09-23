Attribute VB_Name = "modExpedienteServiceFactory"

Option Compare Database
Option Explicit


Public Function CreateExpedienteService(Optional ByVal config As IConfig = Nothing) As IExpedienteService
    On Error GoTo ErrorHandler
    
    Dim effectiveConfig As IConfig
    If config Is Nothing Then
        ' Si no se pasa una configuración, usar la global por defecto
        Set effectiveConfig = modTestContext.GetTestConfig()
    Else
        ' Si se pasa una configuración (desde un test), usarla
        Set effectiveConfig = config
    End If
    
    Dim serviceImpl As New CExpedienteService
    
    Dim repo As IExpedienteRepository
    Set repo = modRepositoryFactory.CreateExpedienteRepository(effectiveConfig)
    
    Dim logger As IOperationLogger
    Set logger = modOperationLoggerFactory.CreateOperationLogger(effectiveConfig)
    
    Dim ErrorHandler As IErrorHandlerService
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService(effectiveConfig)
    
    ' La dependencia de IConfig ahora es manejada por el repositorio, no por el servicio.
    serviceImpl.Initialize repo, logger, ErrorHandler
    
    Set CreateExpedienteService = serviceImpl
    Exit Function
    
ErrorHandler:
    Debug.Print "Error crítico en modExpedienteServiceFactory: " & Err.Description
    Set CreateExpedienteService = Nothing
End Function
