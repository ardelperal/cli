Attribute VB_Name = "TINotificationService"

Option Compare Database
Option Explicit

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE (ESTÁNDAR DE ORO)
' ============================================================================

Public Function TINotificationServiceRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TINotificationService (Estándar de Oro)"
    
    On Error GoTo CleanupSuite
    
    Call SuiteSetup
    suiteResult.AddResult TestSendNotificationSuccessCallsRepositoryCorrectly()
    suiteResult.AddResult TestInitializeWithValidDependencies()
    suiteResult.AddResult TestSendNotificationWithoutInitialize()
    suiteResult.AddResult TestSendNotificationWithInvalidParameters()
    suiteResult.AddResult TestSendNotificationConfigValuesUsed()
    suiteResult.AddResult TestSendNotification_WithCCAndBCC_SavesCorrectly()
    
CleanupSuite:
    
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    
    Set TINotificationServiceRunAll = suiteResult
End Function

' ============================================================================
' PROCEDIMIENTOS HELPER DE LA SUITE
' ============================================================================

Private Sub SuiteSetup()
    Dim db As DAO.Database
    On Error GoTo ErrorHandler
    
    Dim config As IConfig: Set config = modTestContext.GetTestConfig()
    Dim activePath As String: activePath = config.GetValue("CORREOS_DB_PATH")

    Set db = DBEngine.OpenDatabase(activePath, False, False, ";PWD=dpddpd")
    
    
    
    ' BLINDAJE DE IDEMPOTENCIA: Limpiar datos de una ejecución anterior
    db.Execute "DELETE FROM TbCorreosEnviados WHERE Asunto LIKE 'Asunto Test*'", dbFailOnError
    
    db.Close: Set db = Nothing
    
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "TINotificationService.SuiteSetup", Err.Description
End Sub





' ============================================================================
' TESTS INDIVIDUALES
' ============================================================================

Private Function TestSendNotificationSuccessCallsRepositoryCorrectly() As CTestResult
    Set TestSendNotificationSuccessCallsRepositoryCorrectly = New CTestResult
    TestSendNotificationSuccessCallsRepositoryCorrectly.Initialize "SendNotification con éxito debe encolar el correo en la BD"
    
    Dim notificationService As INotificationService
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim localConfig As IConfig
    
    On Error GoTo TestFail
    
    ' ARRANGE: Usar la configuración centralizada del entorno de pruebas
    Set localConfig = modTestContext.GetTestConfig()
    
    ' Crear el servicio real inyectando la configuración local
    Set notificationService = modNotificationServiceFactory.CreateNotificationService(localConfig)
    
    ' ACT: Ejecutar el método a probar
    Dim success As Boolean
    success = notificationService.SendNotification("dest@empresa.com", "Asunto Test", "<html>Cuerpo Test</html>")
    
    ' ASSERT:
    ' 1. Verificar que el método retornó True
    modAssert.AssertTrue success, "SendNotification debe retornar True en caso de éxito."
    ' 2. Verificar que el registro existe REALMENTE en la base de datos
    Dim dbPath As String: dbPath = localConfig.GetValue("CORREOS_DB_PATH")
    Dim dbPassword As String: dbPassword = localConfig.GetValue("CORREOS_PASSWORD")
    Set db = DBEngine.OpenDatabase(dbPath, False, False, ";PWD=" & dbPassword)
    Set rs = db.OpenRecordset("SELECT * FROM TbCorreosEnviados WHERE Asunto = 'Asunto Test'")
    
    modAssert.AssertFalse rs.EOF, "El correo debería haber sido insertado en la tabla TbCorreosEnviados."
    modAssert.AssertEquals "dest@empresa.com", rs!destinatarios.value, "El destinatario en la BD no coincide."
    TestSendNotificationSuccessCallsRepositoryCorrectly.Pass
    GoTo Cleanup

TestFail:
    TestSendNotificationSuccessCallsRepositoryCorrectly.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Set notificationService = Nothing
    Set db = Nothing
    Set rs = Nothing
    Set localConfig = Nothing
End Function

Private Function TestInitializeWithValidDependencies() As CTestResult
    Set TestInitializeWithValidDependencies = New CTestResult
    TestInitializeWithValidDependencies.Initialize "Initialize con dependencias válidas debe tener éxito"
    
    Dim notificationService As INotificationService
    On Error GoTo TestFail
    
    ' Arrange: El servicio obtiene la configuración del contexto centralizado
    ' Act: Intentar crear el servicio usando la factoría
    Set notificationService = modNotificationServiceFactory.CreateNotificationService()
    
    ' Assert
    modAssert.AssertNotNull notificationService, "El servicio no debería ser nulo si las dependencias son válidas."
    TestInitializeWithValidDependencies.Pass
    GoTo Cleanup

TestFail:
    TestInitializeWithValidDependencies.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set notificationService = Nothing
End Function

Private Function TestSendNotificationWithoutInitialize() As CTestResult
    Set TestSendNotificationWithoutInitialize = New CTestResult
    TestSendNotificationWithoutInitialize.Initialize "SendNotification sin inicializar debe fallar devolviendo False"
    
    Dim notificationService As INotificationService
    On Error GoTo TestFail ' Si se produce un error de ejecución, el test falla.
    ' Arrange: Crear la instancia de la clase concreta pero SIN llamar a Initialize
    Dim notificationServiceImpl As New CNotificationService
    Set notificationService = notificationServiceImpl
    
    ' Act: Intentar usar el servicio no inicializado
    Dim success As Boolean
    success = notificationService.SendNotification("test@empresa.com", "Asunto", "<html>Cuerpo</html>")
    
    ' Assert: El servicio debe fallar grácilmente devolviendo False, no con un error.
    modAssert.AssertFalse success, "SendNotification debe devolver False si el servicio no está inicializado."
    TestSendNotificationWithoutInitialize.Pass
    GoTo Cleanup

TestFail:
    TestSendNotificationWithoutInitialize.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set notificationService = Nothing
End Function

Private Function TestSendNotificationWithInvalidParameters() As CTestResult
    Set TestSendNotificationWithInvalidParameters = New CTestResult
    TestSendNotificationWithInvalidParameters.Initialize "SendNotification con parámetros inválidos debe devolver False"
    
    Dim notificationService As INotificationService
    On Error GoTo TestFail
    
    ' Arrange: El servicio obtiene la configuración del contexto centralizado
    Set notificationService = modNotificationServiceFactory.CreateNotificationService()
    
    ' Act & Assert
    modAssert.AssertFalse notificationService.SendNotification("", "Asunto", "Cuerpo"), "Debe devolver False con destinatario vacío."
    modAssert.AssertFalse notificationService.SendNotification("test@test.com", "", "Cuerpo"), "Debe devolver False con asunto vacío."
    modAssert.AssertFalse notificationService.SendNotification("test@test.com", "Asunto", ""), "Debe devolver False con cuerpo vacío."
    TestSendNotificationWithInvalidParameters.Pass
    GoTo Cleanup

TestFail:
    TestSendNotificationWithInvalidParameters.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set notificationService = Nothing
End Function

Private Function TestSendNotificationConfigValuesUsed() As CTestResult
    Set TestSendNotificationConfigValuesUsed = New CTestResult
    TestSendNotificationConfigValuesUsed.Initialize "SendNotification debe usar los valores de configuración correctamente"
    
    Dim notificationService As INotificationService
    Dim db As DAO.Database
    
    On Error GoTo TestFail
    
    ' ARRANGE: Usar la configuración centralizada del entorno de pruebas
    Dim localConfig As IConfig
    Set localConfig = modTestContext.GetTestConfig()
    
    ' Crear el servicio real inyectando la configuración local
    Set notificationService = modNotificationServiceFactory.CreateNotificationService(localConfig)
    ' Act
    Dim success As Boolean
    success = notificationService.SendNotification("dest@empresa.com", "Asunto Test Config", "<html>Test Config</html>")
    
    ' Assert
    modAssert.AssertTrue success, "SendNotification debe funcionar correctamente con config personalizado."
    TestSendNotificationConfigValuesUsed.Pass
    GoTo Cleanup

TestFail:
    TestSendNotificationConfigValuesUsed.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    On Error Resume Next
    Set notificationService = Nothing
    Set localConfig = Nothing
End Function

Private Function TestSendNotification_WithCCAndBCC_SavesCorrectly() As CTestResult
    Set TestSendNotification_WithCCAndBCC_SavesCorrectly = New CTestResult
    TestSendNotification_WithCCAndBCC_SavesCorrectly.Initialize "SendNotification con CC/BCC debe guardar los campos correctos en la BD"
    
    Dim notificationService As INotificationService
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim localConfig As IConfig
    
    On Error GoTo TestFail
    
    ' ARRANGE
    Set localConfig = modTestContext.GetTestConfig()
    Set notificationService = modNotificationServiceFactory.CreateNotificationService(localConfig)

    ' Conectar a la BBDD de prueba para limpiar el estado previo
    Dim dbPath As String: dbPath = localConfig.GetValue("CORREOS_DB_PATH")
    Dim dbPassword As String: dbPassword = localConfig.GetValue("CORREOS_PASSWORD")
    Set db = DBEngine.OpenDatabase(dbPath, False, False, ";PWD=" & dbPassword)
    db.Execute "DELETE FROM TbCorreosEnviados WHERE Asunto = 'Asunto Test CC/BCC'", dbFailOnError

    ' ACT
    Dim success As Boolean
    success = notificationService.SendNotification( _
        destinatarios:="dest@empresa.com", _
        asunto:="Asunto Test CC/BCC", _
        cuerpoHTML:="<html>Cuerpo Test</html>", _
        conCopia:="cc@empresa.com", _
        conCopiaOculta:="bcc@empresa.com")
    
    ' ASSERT
    modAssert.AssertTrue success, "SendNotification con CC/BCC debe retornar True."
    
    Set rs = db.OpenRecordset("SELECT * FROM TbCorreosEnviados WHERE Asunto = 'Asunto Test CC/BCC'")
    
    modAssert.AssertFalse rs.EOF, "El correo con CC/BCC debería haber sido insertado."
    modAssert.AssertEquals "cc@empresa.com", rs!DestinatariosConCopia.value, "El destinatario en CC no coincide."
    modAssert.AssertEquals "bcc@empresa.com", rs!DestinatariosConCopiaOculta.value, "El destinatario en BCC no coincide."
    TestSendNotification_WithCCAndBCC_SavesCorrectly.Pass
    GoTo Cleanup

TestFail:
    TestSendNotification_WithCCAndBCC_SavesCorrectly.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Set notificationService = Nothing
    Set db = Nothing
    Set rs = Nothing
    Set localConfig = Nothing
End Function


