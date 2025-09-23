Attribute VB_Name = "modErrorHandlerFactory"

Option Compare Database
Option Explicit


' =====================================================
' MÓDULO: modErrorHandlerFactory
' DESCRIPCIÓN: Factory para la creación del servicio de errores.
' PATRÓN: CERO ARGUMENTOS (Lección 37)
' =====================================================

Public Function CreateErrorHandlerService(Optional ByVal config As IConfig = Nothing) As IErrorHandlerService
    On Error GoTo ErrorHandler
    
    Dim effectiveConfig As IConfig
    If config Is Nothing Then
        ' Si no se pasa una configuración, usar la global por defecto
        Set effectiveConfig = modTestContext.GetTestConfig()
    Else
        ' Si se pasa una configuración (desde un test), usarla
        Set effectiveConfig = config
    End If
    
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem(effectiveConfig)
    
    Dim errorHandlerImpl As New CErrorHandlerService
    errorHandlerImpl.Initialize effectiveConfig, fs
    
    Set CreateErrorHandlerService = errorHandlerImpl
    Exit Function
    
ErrorHandler:
    Debug.Print "Error crítico en modErrorHandlerFactory.CreateErrorHandlerService: " & Err.Description
    Set CreateErrorHandlerService = Nothing
End Function
