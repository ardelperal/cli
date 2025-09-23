Attribute VB_Name = "modSolicitudServiceFactory"
Option Compare Database
Option Explicit


'******************************************************************************
' MÓDULO: modSolicitudServiceFactory
' DESCRIPCIÓN: Factory para la inyección de dependencias del servicio de solicitudes
' AUTOR: Sistema CONDOR
' FECHA: 2025-01-14
'******************************************************************************

'******************************************************************************
' FACTORY METHODS
'******************************************************************************

'******************************************************************************
' FUNCIÓN: CreateSolicitudService
' DESCRIPCIÓN: Crea una instancia del servicio de solicitudes con todas sus dependencias
' RETORNA: ISolicitudService - Instancia del servicio completamente inicializada
'******************************************************************************
Public Function CreateSolicitudService(Optional ByVal config As IConfig = Nothing) As ISolicitudService
    On Error GoTo ErrorHandler
    
    Dim effectiveConfig As IConfig
    If config Is Nothing Then
        ' Si no se pasa una configuración, usar la global por defecto
        Set effectiveConfig = modTestContext.GetTestConfig()
    Else
        ' Si se pasa una configuración (desde un test), usarla
        Set effectiveConfig = config
    End If
    
    ' 2. Crear dependencias
    Dim errorHandlerService As IErrorHandlerService
    Set errorHandlerService = modErrorHandlerFactory.CreateErrorHandlerService(effectiveConfig)
    
    Dim operationLoggerService As IOperationLogger
    Set operationLoggerService = modOperationLoggerFactory.CreateOperationLogger(effectiveConfig)
    
    ' 3. Crear el repositorio
    Dim solicitudRepo As ISolicitudRepository
    Set solicitudRepo = modRepositoryFactory.CreateSolicitudRepository(effectiveConfig)
    
    ' 4. Crear el servicio de autenticación
    Dim authService As IAuthService
    Set authService = modAuthFactory.CreateAuthService(effectiveConfig)
    
    ' 5. Crear el servicio de workflow
    Dim workflowService As IWorkflowService
    Set workflowService = modWorkflowServiceFactory.CreateWorkflowService(effectiveConfig)
    
    ' 6. Crear e inicializar la instancia del servicio
    Dim serviceInstance As New CSolicitudService
    serviceInstance.Initialize solicitudRepo, operationLoggerService, errorHandlerService, authService, workflowService
    
    ' 6. Devolver la instancia como el tipo de la interfaz
    Set CreateSolicitudService = serviceInstance
    
    Exit Function
    
ErrorHandler:
    ' Usar Debug.Print en una factoría es aceptable si errorHandler falla.
    Debug.Print "Error crítico en modSolicitudServiceFactory.CreateSolicitudService: " & Err.Description
    Set CreateSolicitudService = Nothing
End Function
