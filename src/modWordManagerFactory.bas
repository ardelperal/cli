Attribute VB_Name = "modWordManagerFactory"
Option Compare Database
Option Explicit



' =====================================================
' MÓDULO: modWordManagerFactory
' DESCRIPCIÓN: Factory para la creación del servicio de gestión de Word
' AUTOR: Sistema CONDOR
' FECHA: 2025-01-15
' =====================================================

Public Function CreateWordManager(Optional ByVal config As IConfig = Nothing) As IWordManager
    On Error GoTo ErrorHandler
    
    Dim effectiveConfig As IConfig
    If config Is Nothing Then
        ' Si no se pasa una configuración, usar la global por defecto
        Set effectiveConfig = modTestContext.GetTestConfig()
    Else
        ' Si se pasa una configuración (desde un test), usarla
        Set effectiveConfig = config
    End If
    
    Dim wordApp As Object
    Dim ErrorHandler As IErrorHandlerService
    
    ' Crear dependencias
    Set ErrorHandler = modErrorHandlerFactory.CreateErrorHandlerService(effectiveConfig)
    
    ' Crear instancia de Word y luego inicializar
    Set wordApp = CreateObject("Word.Application")
    wordApp.Visible = False
    wordApp.DisplayAlerts = False
    
    Dim wordManagerInstance As New CWordManager
    wordManagerInstance.Initialize wordApp, ErrorHandler
    Set CreateWordManager = wordManagerInstance
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error en modWordManagerFactory.CreateWordManager: " & Err.Description
    Err.Raise Err.Number, "modWordManagerFactory.CreateWordManager", Err.Description
End Function
