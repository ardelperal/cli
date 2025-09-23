Attribute VB_Name = "ModAppManagerFactory"
Option Compare Database
Option Explicit
'CAMBIOS DE ANDRÉS

Public Function CreateAppManager() As IAppManager
    On Error GoTo ErrorHandler
    
    Dim appManagerImpl As New CAppManager
    
    ' Crear dependencias usando sus respectivas factorías
    Dim authSvc As IAuthService
    Set authSvc = modAuthFactory.CreateAuthService()
    
    Dim configSvc As IConfig
    Set configSvc = modConfigFactory.CreateConfigService()
    
    Dim ErrorHandler As IErrorHandlerService
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService()
    
    ' Inyectar dependencias
    appManagerImpl.Initialize authSvc, configSvc, ErrorHandler
    
    Set CreateAppManager = appManagerImpl
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error fatal en ModAppManagerFactory.CreateAppManager: " & Err.Description
    Set CreateAppManager = Nothing
End Function
