VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockOperationLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IOperationLogger

Private m_SaveLog_WasCalled As Boolean
Private m_SaveLog_LastEntry As EOperationLog
Private m_LogSolicitudOperation_WasCalled As Boolean
Private m_LogSolicitudOperation_LastEntry As EOperationLog
Private m_LogSolicitudOperation_LastSolicitud As ESolicitud
Private m_LogSolicitudOperation_LastUserId As String

Public Property Get SaveLog_WasCalled() As Boolean: SaveLog_WasCalled = m_SaveLog_WasCalled: End Property
Public Property Get SaveLog_LastEntry() As EOperationLog: Set SaveLog_LastEntry = m_SaveLog_LastEntry: End Property
Public Property Get LogSolicitudOperation_WasCalled() As Boolean: LogSolicitudOperation_WasCalled = m_LogSolicitudOperation_WasCalled: End Property
Public Property Get LogSolicitudOperation_LastEntry() As EOperationLog: Set LogSolicitudOperation_LastEntry = m_LogSolicitudOperation_LastEntry: End Property
Public Property Get LogSolicitudOperation_LastSolicitud() As ESolicitud: Set LogSolicitudOperation_LastSolicitud = m_LogSolicitudOperation_LastSolicitud: End Property
Public Property Get LogSolicitudOperation_LastUserId() As String: LogSolicitudOperation_LastUserId = m_LogSolicitudOperation_LastUserId: End Property

Private Sub Class_Initialize(): Call Reset: End Sub

Public Sub Reset()
    m_SaveLog_WasCalled = False
    Set m_SaveLog_LastEntry = Nothing
    
    m_LogSolicitudOperation_WasCalled = False
    Set m_LogSolicitudOperation_LastEntry = Nothing
    Set m_LogSolicitudOperation_LastSolicitud = Nothing
    m_LogSolicitudOperation_LastUserId = ""
End Sub

Private Sub IOperationLogger_LogOperation(ByVal logEntry As EOperationLog)
    m_SaveLog_WasCalled = True
    Set m_SaveLog_LastEntry = logEntry
End Sub

Private Sub IOperationLogger_LogSolicitudOperation(ByVal logEntry As EOperationLog, ByVal solicitud As ESolicitud, ByVal userId As String)
    m_LogSolicitudOperation_WasCalled = True
    Set m_LogSolicitudOperation_LastEntry = logEntry
    Set m_LogSolicitudOperation_LastSolicitud = solicitud
    m_LogSolicitudOperation_LastUserId = userId
End Sub
