VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CNotificationRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements INotificationRepository

Private m_configService As IConfig
Private m_ErrorHandler As IErrorHandlerService
Private m_isInitialized As Boolean

Public Sub Initialize(ByVal configSrv As IConfig, ByVal errorSrv As IErrorHandlerService)
    Set m_configService = configSrv
    Set m_ErrorHandler = errorSrv
    m_isInitialized = True
End Sub

Private Function INotificationRepository_EnqueueEmail(ByVal destinatarios As String, ByVal asunto As String, ByVal cuerpoHTML As String, Optional ByVal conCopia As String = "", Optional ByVal conCopiaOculta As String = "", Optional ByVal adjuntoPath As String = "") As Boolean
    On Error GoTo ErrorHandler
    
    INotificationRepository_EnqueueEmail = False
    If Not m_isInitialized Then Exit Function
    
    Dim db As DAO.Database, rs As DAO.Recordset, rsMax As DAO.Recordset
    Dim correosDbPath As String, dbPassword As String
    
    correosDbPath = m_configService.GetCorreosDBPath()
    dbPassword = m_configService.GetCorreosPassword()
    
    If Len(correosDbPath) = 0 Then Exit Function
    
    If Len(dbPassword) > 0 Then
        Set db = DBEngine.OpenDatabase(correosDbPath, False, False, ";PWD=" & dbPassword)
    Else
        Set db = DBEngine.OpenDatabase(correosDbPath, False, False)
    End If
    
    ' PASO 1: Calcular el nuevo ID manualmente usando una consulta SQL expl?cita
    Dim nuevoId As Long
    Dim maxId As Variant
    
    Set rsMax = db.OpenRecordset("SELECT MAX(IDCorreo) AS MaxID FROM TbCorreosEnviados", dbOpenSnapshot)
    If Not rsMax.EOF Then
        maxId = rsMax.Fields("MaxID").value
    End If
    rsMax.Close
    
    nuevoId = Nz(maxId, 0) + 1
    
    ' PASO 2: Insertar el nuevo registro con el ID calculado
    Set rs = db.OpenRecordset("TbCorreosEnviados", dbOpenDynaset)
    
    rs.AddNew
    rs("IDCorreo") = nuevoId ' Asignar el nuevo ID
    rs("Destinatarios") = destinatarios
    rs("Asunto") = asunto
    rs("Cuerpo") = cuerpoHTML
    rs("DestinatariosConCopia") = conCopia
    rs("DestinatariosConCopiaOculta") = conCopiaOculta
    rs("URLAdjunto") = adjuntoPath
    rs("FechaGrabacion") = Now()
    rs.Update
    
    rs.Close
    db.Close
    
    INotificationRepository_EnqueueEmail = True
    Exit Function
    
ErrorHandler:
    If Not rsMax Is Nothing Then rsMax.Close
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    If Not m_ErrorHandler Is Nothing Then
        m_ErrorHandler.LogError Err.Number, Err.Description, "CNotificationRepository.EnqueueEmail"
    End If
End Function
