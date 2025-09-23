VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COperationLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IOperationLogger

Private m_configService As IConfig
Private m_OperationRepository As IOperationRepository
Private m_ErrorHandler As IErrorHandlerService

Public Sub Initialize(config As IConfig, repository As IOperationRepository, ErrorHandler As IErrorHandlerService)
    Set m_configService = config
    Set m_OperationRepository = repository
    Set m_ErrorHandler = ErrorHandler
End Sub

Private Sub IOperationLogger_LogOperation(ByVal logEntry As EOperationLog)
    On Error GoTo ErrorHandler
    ' Simplemente delega la operaci?n de guardado al repositorio
    m_OperationRepository.SaveLog logEntry
    Exit Sub
ErrorHandler:
    m_ErrorHandler.LogError Err.Number, Err.Description, "COperationLogger.LogOperation"
End Sub

Private Sub IOperationLogger_LogSolicitudOperation(ByVal logEntry As EOperationLog, ByVal solicitud As ESolicitud, ByVal userId As String)
    On Error GoTo ErrorHandler
    ' Enriquece el objeto log con datos adicionales antes de guardarlo
    logEntry.usuario = userId
    logEntry.idEntidadAfectada = solicitud.idSolicitud
    ' Delega la operaci?n de guardado al repositorio
    m_OperationRepository.SaveLog logEntry
    Exit Sub
ErrorHandler:
    m_ErrorHandler.LogError Err.Number, Err.Description, "COperationLogger.LogSolicitudOperation"
End Sub

' --- M?todos P?blicos de Conveniencia ---

Public Sub LogOperation(ByVal logEntry As EOperationLog)
    Call IOperationLogger_LogOperation(logEntry)
End Sub

Public Sub LogSolicitudOperation(ByVal logEntry As EOperationLog, ByVal solicitud As ESolicitud, ByVal userId As String)
    Call IOperationLogger_LogSolicitudOperation(logEntry, solicitud, userId)
End Sub
