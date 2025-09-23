Attribute VB_Name = "modDocumentServiceFactory"
Option Compare Database
Option Explicit


Public Function CreateDocumentService(Optional ByVal config As IConfig = Nothing) As IDocumentService
    On Error GoTo ErrorHandler
    
    Dim effectiveConfig As IConfig
    If config Is Nothing Then
        ' Si no se pasa una configuración, usar la global por defecto
        Set effectiveConfig = modTestContext.GetTestConfig()
    Else
        ' Si se pasa una configuración (desde un test), usarla
        Set effectiveConfig = config
    End If
    
    Dim serviceImpl As New CDocumentService
    
    ' Crear TODAS las dependencias
    Dim wordMgr As IWordManager
    Set wordMgr = modWordManagerFactory.CreateWordManager(effectiveConfig)
    
    Dim errHandler As IErrorHandlerService
    Set errHandler = modErrorHandlerFactory.CreateErrorHandlerService(effectiveConfig)
    
    Dim solicitudSrv As ISolicitudService
    Set solicitudSrv = modSolicitudServiceFactory.CreateSolicitudService(effectiveConfig)
    
    Dim mapeoRepo As IMapeoRepository
    Set mapeoRepo = modRepositoryFactory.CreateMapeoRepository(effectiveConfig)
    
    Dim fileSystem As IFileSystem
    Set fileSystem = modFileSystemFactory.CreateFileSystem(effectiveConfig)
    
    ' Inyectar dependencias en el orden correcto
    serviceImpl.Initialize effectiveConfig, fileSystem, wordMgr, errHandler, solicitudSrv, mapeoRepo
    
    Set CreateDocumentService = serviceImpl
    Exit Function
    
ErrorHandler:
    Debug.Print "Error crítico en modDocumentServiceFactory: " & Err.Description
    Set CreateDocumentService = Nothing
End Function
