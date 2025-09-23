VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWorkflowRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IWorkflowRepository

Private m_Config As IConfig
Private m_ErrorHandler As IErrorHandlerService

Public Sub Initialize(ByVal config As IConfig, ByVal ErrorHandler As IErrorHandlerService)
    Set m_Config = config
    Set m_ErrorHandler = ErrorHandler
End Sub

Private Function IWorkflowRepository_IsValidTransition(ByVal tipoSolicitud As String, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal usuarioRol As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim isValid As Boolean
    isValid = False
    
    Set db = DBEngine.OpenDatabase(m_Config.GetCondorDataPath(), False, False, ";PWD=" & m_Config.GetCondorPassword())
    
    Dim strSQL As String
    strSQL = "SELECT Count(T.idTransicion) As TransicionExiste " & _
             "FROM (tbTransiciones AS T INNER JOIN tbEstados AS EO ON T.idEstadoOrigen = EO.idEstado) " & _
             "INNER JOIN tbEstados AS ED ON T.idEstadoDestino = ED.idEstado " & _
             "WHERE EO.nombreEstado = [prmEstadoOrigen] AND ED.nombreEstado = [prmEstadoDestino] AND T.rolRequerido = [prmRolRequerido];"
             
    Set qdf = db.CreateQueryDef("", strSQL)
    qdf.Parameters("prmEstadoOrigen") = estadoOrigen
    qdf.Parameters("prmEstadoDestino") = estadoDestino
    qdf.Parameters("prmRolRequerido") = usuarioRol
    
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    If Not rs.EOF Then
        isValid = (rs!TransicionExiste.value > 0)
    End If
    
    IWorkflowRepository_IsValidTransition = isValid

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then qdf.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing
    Set qdf = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    m_ErrorHandler.LogError Err.Number, "Error en CWorkflowRepository.IsValidTransition: " & Err.Description, "CWorkflowRepository"
    IWorkflowRepository_IsValidTransition = False
    GoTo Cleanup
End Function

Private Function IWorkflowRepository_GetNextStates(ByVal idEstadoActual As Long, ByVal usuarioRol As String) As Object
    ' Implementaci?n real para obtener los siguientes estados v?lidos desde la BD.
    On Error GoTo ErrorHandler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim dict As New Scripting.Dictionary
    
    ' Obtener la conexi?n a la base de datos del backend
    Set db = DBEngine.OpenDatabase(m_Config.GetCondorDataPath(), False, False, ";PWD=" & m_Config.GetCondorPassword())
    
    ' Consulta SQL parametrizada para seguridad y claridad (Lecci?n 18)
    ' Une transiciones y estados para obtener los nombres de los estados de destino.
    ' Actualizada para trabajar con los nuevos 7 estados del flujo refactorizado
    Dim strSQL As String
    strSQL = "SELECT T.idEstadoDestino, E.nombreEstado " & _
             "FROM tbTransiciones AS T " & _
             "INNER JOIN tbEstados AS E ON T.idEstadoDestino = E.idEstado " & _
             "WHERE T.idEstadoOrigen = [prmIdEstadoActual] AND T.rolRequerido = [prmRolRequerido];"
             
    Set qdf = db.CreateQueryDef("", strSQL)
    qdf.Parameters("prmIdEstadoActual") = idEstadoActual
    qdf.Parameters("prmRolRequerido") = usuarioRol
    
    Set rs = qdf.OpenRecordset(dbOpenSnapshot)
    
    ' Poblar el diccionario con los resultados (ID como clave, Nombre como valor)
    While Not rs.EOF
        dict.Add key:=rs!IdEstadoDestino.value, Item:=rs!NombreEstado.value
        rs.MoveNext
    Wend
    
    Set IWorkflowRepository_GetNextStates = dict

Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then qdf.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing
    Set qdf = Nothing
    Set db = Nothing
    Exit Function

ErrorHandler:
    m_ErrorHandler.LogError Err.Number, "Error en CWorkflowRepository.GetNextStates: " & Err.Description, "CWorkflowRepository"
    Set IWorkflowRepository_GetNextStates = dict ' Devolver diccionario vac?o en caso de error
    GoTo Cleanup
End Function

