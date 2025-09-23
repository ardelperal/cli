Attribute VB_Name = "TestSolicitudService"
Option Compare Database
Option Explicit


' ============================================================================
' SUITE DE PRUEBAS UNITARIAS PARA CSolicitudService
' Arquitectura: Pruebas Aisladas con Mocks
' ============================================================================

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE
' ============================================================================

Public Function TestSolicitudServiceRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    Call suiteResult.Initialize("TestSolicitudService - Pruebas Unitarias CSolicitudService")
    
    Call suiteResult.AddResult(TestCreateSolicitudSuccess())
    Call suiteResult.AddResult(TestCreateSolicitudFailsWithEmptyExpediente())
    Call suiteResult.AddResult(TestSaveSolicitudSuccess())
    
    Set TestSolicitudServiceRunAll = suiteResult
End Function



' ============================================================================
' PRUEBAS
' ============================================================================

Private Function TestCreateSolicitudSuccess() As CTestResult
    Set TestCreateSolicitudSuccess = New CTestResult
    TestCreateSolicitudSuccess.Initialize "CreateSolicitud debe crear una solicitud con valores por defecto correctos"
    
    Dim serviceImpl As CSolicitudService
    Dim mockRepo As CMockSolicitudRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim mockAuth As CMockAuthService
    Dim mockWorkflow As CMockWorkflowService
    Dim service As ISolicitudService
    Dim expediente As EExpediente
    Dim result As ESolicitud
    
    On Error GoTo TestFail
    
    ' Arrange
    Set mockRepo = New CMockSolicitudRepository
    Set mockLogger = New CMockOperationLogger
    Set mockErrorHandler = New CMockErrorHandlerService
    Set mockAuth = New CMockAuthService
    Set mockWorkflow = New CMockWorkflowService
    mockAuth.ConfigureGetCurrentUserEmail "creator.user@condor.com"
    
    Set serviceImpl = New CSolicitudService
    serviceImpl.Initialize mockRepo, mockLogger, mockErrorHandler, mockAuth, mockWorkflow
    Set service = serviceImpl
    
    Set expediente = New EExpediente
    expediente.idExpediente = 123
    
    mockRepo.ConfigureSaveSolicitud 456 ' Simular que el guardado devuelve un nuevo ID
    
    ' Act
    Set result = service.CreateSolicitud(expediente)
    
    ' Assert
    modAssert.AssertNotNull result, "La solicitud devuelta no debe ser nula."
    modAssert.AssertTrue mockRepo.SaveSolicitudCalled, "Se debe llamar al método SaveSolicitud del repositorio."
    modAssert.AssertEquals "creator.user@condor.com", mockRepo.SaveSolicitud_LastSolicitud.usuarioCreacion, "El usuario de creación debe ser establecido por el servicio de autenticación."
    
    TestCreateSolicitudSuccess.Pass
    GoTo Cleanup
    
TestFail:
    TestCreateSolicitudSuccess.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set serviceImpl = Nothing
    Set mockRepo = Nothing
    Set mockLogger = Nothing
    Set mockErrorHandler = Nothing
    Set mockAuth = Nothing
    Set mockWorkflow = Nothing
    Set service = Nothing
    Set expediente = Nothing
    Set result = Nothing
End Function

Private Function TestCreateSolicitudFailsWithEmptyExpediente() As CTestResult
    Set TestCreateSolicitudFailsWithEmptyExpediente = New CTestResult
    TestCreateSolicitudFailsWithEmptyExpediente.Initialize "CreateSolicitud debe fallar si idExpediente está vacío"
    
    Dim service As ISolicitudService
    Dim expediente As EExpediente
    On Error GoTo TestExpectedFail
    
    ' Arrange
    Set service = modSolicitudServiceFactory.CreateSolicitudService() ' Usar factory para dependencias
    Set expediente = New EExpediente
    expediente.idExpediente = 0 ' Usar 0 en lugar de "" para un Long
    
    ' Act
    service.CreateSolicitud expediente
    
    ' Assert - Si llegamos aquí, la prueba ha fallado
    TestCreateSolicitudFailsWithEmptyExpediente.Fail "La función debería haber lanzado un error."
    GoTo Cleanup
    
TestExpectedFail:
    ' El error es esperado, la prueba ha pasado.
    TestCreateSolicitudFailsWithEmptyExpediente.Pass
    
Cleanup:
    Set service = Nothing
    Set expediente = Nothing
End Function

Private Function TestSaveSolicitudSuccess() As CTestResult
    Set TestSaveSolicitudSuccess = New CTestResult
    TestSaveSolicitudSuccess.Initialize "SaveSolicitud debe establecer los campos de modificación"
    
    ' Variables locales
    Dim serviceImpl As CSolicitudService
    Dim mockRepo As CMockSolicitudRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim mockAuth As CMockAuthService
    Dim mockWorkflow As CMockWorkflowService
    Dim service As ISolicitudService ' Variable de interfaz para el test
    Dim solicitud As ESolicitud
    
    On Error GoTo TestFail
    
    ' Arrange
    Set mockRepo = New CMockSolicitudRepository
    Set mockLogger = New CMockOperationLogger
    Set mockErrorHandler = New CMockErrorHandlerService
    Set mockAuth = New CMockAuthService
    Set mockWorkflow = New CMockWorkflowService
    mockAuth.ConfigureGetCurrentUserEmail "test.user@condor.com"

    ' PATRÓN CORRECTO: Instanciar la clase concreta, inicializarla, y LUEGO asignarla a la interfaz
    Set serviceImpl = New CSolicitudService
    serviceImpl.Initialize mockRepo, mockLogger, mockErrorHandler, mockAuth, mockWorkflow
    Set service = serviceImpl
    
    Set solicitud = New ESolicitud
    solicitud.idSolicitud = 456
    solicitud.idEstadoInterno = 1 ' Estado Borrador
    
    ' Act
    Call service.SaveSolicitud(solicitud)
    
    ' Assert
    modAssert.AssertTrue mockRepo.SaveSolicitudCalled, "Se debe llamar al método SaveSolicitud del repositorio"
    modAssert.AssertEquals "test.user@condor.com", mockRepo.SaveSolicitud_LastSolicitud.usuarioModificacion, "El campo usuarioModificacion debe ser establecido correctamente"
    
    TestSaveSolicitudSuccess.Pass
    GoTo Cleanup
    
TestFail:
    TestSaveSolicitudSuccess.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set serviceImpl = Nothing
    Set mockRepo = Nothing
    Set mockLogger = Nothing
    Set mockErrorHandler = Nothing
    Set mockAuth = Nothing
    Set mockWorkflow = Nothing
    Set service = Nothing
    Set solicitud = Nothing
End Function
