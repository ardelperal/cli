Attribute VB_Name = "modNotificationServiceFactory"

Option Compare Database
Option Explicit



' =====================================================
' MÓDULO: modNotificationServiceFactory
' DESCRIPCIÓN: Factory especializada para la creación del servicio de notificaciones
' AUTOR: Sistema CONDOR
' FECHA: 2024
' =====================================================

' Función factory para crear y configurar el servicio de notificaciones
Public Function CreateNotificationService(Optional ByVal config As IConfig = Nothing) As INotificationService
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
    
    ' Crear el repositorio usando el factory
    Dim notificationRepository As INotificationRepository
    Set notificationRepository = modRepositoryFactory.CreateNotificationRepository(effectiveConfig)
    
    ' Crear una instancia de la clase concreta
    Dim notificationServiceInstance As New CNotificationService
    
    ' Inicializar la instancia concreta con todas las dependencias
    Call notificationServiceInstance.Initialize(effectiveConfig, operationLogger, notificationRepository, ErrorHandler)
    
    ' Devolver la instancia inicializada como el tipo de la interfaz
    Set CreateNotificationService = notificationServiceInstance
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error en modNotificationServiceFactory.CreateNotificationService: " & Err.Number & " - " & Err.Description
    Err.Raise Err.Number, Err.Source, Err.Description
End Function
