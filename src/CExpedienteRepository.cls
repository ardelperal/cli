VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExpedienteRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IExpedienteRepository

Private Const MODULE_NAME As String = "CExpedienteRepository"

Private m_Config As IConfig
Private m_ErrorHandler As IErrorHandlerService

Public Sub Initialize(ByVal config As IConfig, ByVal ErrorHandler As IErrorHandlerService)
    Set m_Config = config
    Set m_ErrorHandler = ErrorHandler
End Sub

' --- IMPLEMENTACI?N DE LA INTERFAZ (PRIVADA) ---

Private Function IExpedienteRepository_ObtenerExpedientePorId(ByVal idExpediente As Long) As EExpediente
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim expediente As EExpediente
    
    On Error GoTo ErrorHandler
    
    Set db = GetDatabaseConnection()
    
    Set qdf = db.CreateQueryDef("", "SELECT * FROM TbExpedientes WHERE IDExpediente = [pIdExpediente]")
    qdf.Parameters("[pIdExpediente]").value = idExpediente
    
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    If Not rs.EOF Then
        Set expediente = MapRecordsetToExpediente(rs)
    End If
    
    Set IExpedienteRepository_ObtenerExpedientePorId = expediente

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then qdf.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing
    Set qdf = Nothing
    Set db = Nothing
    Set expediente = Nothing
    Exit Function

ErrorHandler:
    m_ErrorHandler.LogError Err.Number, "Error en " & MODULE_NAME & ".IExpedienteRepository_ObtenerExpedientePorId: " & Err.Description
    Resume Cleanup
End Function

Private Function IExpedienteRepository_ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As EExpediente
    ' TODO: Implementar l?gica de base de datos
    Set IExpedienteRepository_ObtenerExpedientePorNemotecnico = New EExpediente
End Function

Private Function IExpedienteRepository_ObtenerExpedientesActivosParaSelector() As Object
    ' TODO: Implementar l?gica de base de datos
    Set IExpedienteRepository_ObtenerExpedientesActivosParaSelector = New Scripting.Dictionary
End Function

' --- M?TODOS P?BLICOS DE CONVENIENCIA (LECCI?N 24) ---

Public Function ObtenerExpedientePorId(ByVal idExpediente As Long) As EExpediente
    Set ObtenerExpedientePorId = IExpedienteRepository_ObtenerExpedientePorId(idExpediente)
End Function

Public Function ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As EExpediente
    Set ObtenerExpedientePorNemotecnico = IExpedienteRepository_ObtenerExpedientePorNemotecnico(Nemotecnico)
End Function

Public Function ObtenerExpedientesActivosParaSelector() As Object
    Set ObtenerExpedientesActivosParaSelector = IExpedienteRepository_ObtenerExpedientesActivosParaSelector()
End Function

' --- FUNCIONES AUXILIARES (PRIVADAS) ---

Private Function GetDatabaseConnection() As DAO.Database
    On Error GoTo ErrorHandler
    
    Dim dbPath As String
    dbPath = m_Config.GetExpedientesDataPath()
    
    Dim dbPassword As String
    dbPassword = m_Config.GetExpedientesPassword()
    
    If Dir(dbPath) = "" Then
        Err.Raise vbObjectError + 1001, MODULE_NAME, "La base de datos de expedientes no se encontr? en la ruta: " & dbPath
    End If
    
    Set GetDatabaseConnection = DBEngine.OpenDatabase(dbPath, False, True, ";PWD=" & dbPassword)
    Exit Function
    
ErrorHandler:
    m_ErrorHandler.LogError Err.Number, "Error en " & MODULE_NAME & ".GetDatabaseConnection: " & Err.Description
    Set GetDatabaseConnection = Nothing
End Function

Private Function MapRecordsetToExpediente(ByVal rs As DAO.Recordset) As EExpediente
    On Error GoTo ErrorHandler
    
    Dim expediente As New EExpediente
    
    ' Mapeo completo y seguro de propiedades
    expediente.idExpediente = Nz(rs!idExpediente.value, 0)
    expediente.Nemotecnico = Nz(rs!Nemotecnico.value, "")
    expediente.Titulo = Nz(rs!Titulo.value, "")
    expediente.ContratistaPrincipal = Nz(rs!ContratistaPrincipal.value, "")
    expediente.Pecal = Nz(rs!Pecal.value, "")
    
    Set MapRecordsetToExpediente = expediente
    
    Exit Function
ErrorHandler:
    m_ErrorHandler.LogError Err.Number, "Error en " & MODULE_NAME & ".MapRecordsetToExpediente: " & Err.Description
    Set MapRecordsetToExpediente = Nothing
End Function
