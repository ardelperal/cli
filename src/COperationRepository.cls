VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COperationRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IOperationRepository

Private m_configService As IConfig
Private m_ErrorHandler As IErrorHandlerService

Public Sub Initialize(config As IConfig, ErrorHandler As IErrorHandlerService)
    Set m_configService = config
    Set m_ErrorHandler = ErrorHandler
End Sub

Private Sub IOperationRepository_SaveLog(ByVal logEntry As EOperationLog)
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    
    Set db = DBEngine.OpenDatabase(m_configService.GetCondorDataPath(), False, False, ";PWD=" & m_configService.GetCondorPassword())
    
    Set qdf = db.CreateQueryDef("", INSERT_OPERATION_LOG)
    
    qdf.Parameters("pUsuario") = logEntry.usuario
    qdf.Parameters("pTipoOperacion") = logEntry.tipoOperacion
    qdf.Parameters("pEntidad") = logEntry.entidad
    qdf.Parameters("pIdEntidad") = logEntry.idEntidadAfectada
    qdf.Parameters("pDescripcion") = logEntry.descripcion
    qdf.Parameters("pResultado") = logEntry.resultado
    qdf.Parameters("pDetalles") = logEntry.detalles
    
    qdf.Execute dbFailOnError
    
    Set qdf = Nothing
    db.Close
    Exit Sub
    
ErrorHandler:
    m_ErrorHandler.LogError Err.Number, Err.Description, "COperationRepository.SaveLog"
    If Not qdf Is Nothing Then Set qdf = Nothing
    If Not db Is Nothing Then db.Close
End Sub

Public Sub SaveLog(ByVal logEntry As EOperationLog)
    Call IOperationRepository_SaveLog(logEntry)
End Sub
