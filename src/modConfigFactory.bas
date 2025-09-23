Attribute VB_Name = "modConfigFactory"
Option Compare Database
Option Explicit

' =====================================================
' MÓDULO: modConfigFactory
' DESCRIPCIÓN: Factory para el servicio de configuración.
' ARQUITECTURA: Autónomo, sin dependencias externas para
' evitar ciclos de recursión.
' =====================================================

Public Function CreateConfigService() As IConfig
    On Error GoTo FactoryErrorHandler
    
    Dim configImpl As New CConfig
    configImpl.LoadConfiguration ' La configuración se carga al crear
    
    Set CreateConfigService = configImpl
    Exit Function
    
FactoryErrorHandler:
    ' ESTA FACTORÍA ES DE NIVEL 0. NO PUEDE DEPENDER DEL ERRORHANDLER.
    ' Si falla, debe notificar directamente y detener la ejecución.
    ' Error crítico - no mostrar diálogos en CLI
    On Error Resume Next
    Err.Clear
    Set CreateConfigService = Nothing
End Function


