Attribute VB_Name = "modWorkflowServiceFactory"
Option Compare Database
Option Explicit


Public Function CreateWorkflowService(Optional ByVal config As IConfig = Nothing) As IWorkflowService
    On Error GoTo ErrorHandler

    Dim effectiveConfig As IConfig
    If config Is Nothing Then
        ' Si no se pasa una configuración, usar la global por defecto
        Set effectiveConfig = modTestContext.GetTestConfig()
    Else
        ' Si se pasa una configuración (desde un test), usarla
        Set effectiveConfig = config
    End If

    ' 2. Crear las dependencias
    Dim repo As IWorkflowRepository
    Set repo = modRepositoryFactory.CreateWorkflowRepository(effectiveConfig)

    Dim logger As IOperationLogger
    Set logger = modOperationLoggerFactory.CreateOperationLogger(effectiveConfig)

    Dim errorHandlerSvc As IErrorHandlerService
    Set errorHandlerSvc = modErrorHandlerFactory.CreateErrorHandlerService(effectiveConfig)

    ' 3. Crear e inicializar la implementación del servicio
    Dim serviceImpl As New CWorkflowService
    serviceImpl.Initialize repo, logger, errorHandlerSvc

    Set CreateWorkflowService = serviceImpl
    Exit Function

ErrorHandler:
    Debug.Print "Error crítico en modWorkflowServiceFactory: " & Err.Description
    Set CreateWorkflowService = Nothing
End Function

Public Function CreateWorkflowServiceWithMocks(ByVal mockRepo As IWorkflowRepository) As IWorkflowService
    Dim serviceImpl As New CWorkflowService
    ' Inyectar el mock del repositorio y mocks pasivos para las otras dependencias
    serviceImpl.Initialize mockRepo, New CMockOperationLogger, New CMockErrorHandlerService
    Set CreateWorkflowServiceWithMocks = serviceImpl
End Function
