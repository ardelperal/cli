Attribute VB_Name = "TestOperationLogger"
Option Compare Database
Option Explicit

Public Function TestOperationLoggerRunAll() As CTestSuiteResult
    Set TestOperationLoggerRunAll = New CTestSuiteResult
    TestOperationLoggerRunAll.Initialize "TestOperationLogger"
    TestOperationLoggerRunAll.AddResult TestLogOperation_CallsRepository()
    TestOperationLoggerRunAll.AddResult TestLogSolicitudOperation_EnrichesLogEntry()
End Function

Private Function TestLogOperation_CallsRepository() As CTestResult
    Set TestLogOperation_CallsRepository = New CTestResult
    TestLogOperation_CallsRepository.Initialize "LogOperation debe llamar al método SaveLog del repositorio"

    Dim serviceImpl As COperationLogger
    Dim mockRepo As CMockOperationRepository
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim service As IOperationLogger
    Dim logEntry As EOperationLog
    
    On Error GoTo TestFail

    ' Arrange
    Set mockRepo = New CMockOperationRepository
    Set mockErrorHandler = New CMockErrorHandlerService
    
    Set serviceImpl = New COperationLogger
    serviceImpl.Initialize Nothing, mockRepo, mockErrorHandler
    Set service = serviceImpl
    
    Set logEntry = New EOperationLog
    logEntry.Initialize Now, "test_user", "TEST", "TestEntity", 1, "Desc", "OK", "Details"
    
    ' Act
    service.LogOperation logEntry
    
    ' Assert
    modAssert.AssertTrue mockRepo.SaveLog_WasCalled, "El método SaveLog del repositorio debería haber sido llamado."
    modAssert.AssertEquals "TEST", mockRepo.SaveLog_LastEntry.tipoOperacion, "El tipo de operación no coincide."

    TestLogOperation_CallsRepository.Pass
    GoTo Cleanup

TestFail:
    TestLogOperation_CallsRepository.Fail "Error: " & Err.Description
    
Cleanup:
    Set serviceImpl = Nothing
    Set mockRepo = Nothing
    Set mockErrorHandler = Nothing
    Set service = Nothing
    Set logEntry = Nothing
End Function

Private Function TestLogSolicitudOperation_EnrichesLogEntry() As CTestResult
    Set TestLogSolicitudOperation_EnrichesLogEntry = New CTestResult
    TestLogSolicitudOperation_EnrichesLogEntry.Initialize "LogSolicitudOperation debe enriquecer el log y llamar al repositorio"

    Dim serviceImpl As COperationLogger
    Dim mockRepo As CMockOperationRepository
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim service As IOperationLogger
    Dim logEntry As EOperationLog
    Dim mockSolicitud As ESolicitud
    Dim testUserId As String
    
    On Error GoTo TestFail

    ' Arrange
    Set mockRepo = New CMockOperationRepository
    Set mockErrorHandler = New CMockErrorHandlerService
    
    Set serviceImpl = New COperationLogger
    serviceImpl.Initialize Nothing, mockRepo, mockErrorHandler
    Set service = serviceImpl
    
    Set logEntry = New EOperationLog
    Set mockSolicitud = New ESolicitud
    mockSolicitud.idSolicitud = 123 ' ID de prueba
    testUserId = "test.user@example.com"
    
    ' Act
    service.LogSolicitudOperation logEntry, mockSolicitud, testUserId
    
    ' Assert
    modAssert.AssertTrue mockRepo.SaveLog_WasCalled, "El método SaveLog del repositorio debería haber sido llamado."
    
    Dim capturedLog As EOperationLog
    Set capturedLog = mockRepo.SaveLog_LastEntry
    
    modAssert.AssertNotNull capturedLog, "El objeto log capturado no debería ser nulo."
    modAssert.AssertEquals testUserId, capturedLog.usuario, "El UserID no fue enriquecido correctamente en el log."
    modAssert.AssertEquals mockSolicitud.idSolicitud, capturedLog.idEntidadAfectada, "El ID de la solicitud no fue enriquecido correctamente."
    
    TestLogSolicitudOperation_EnrichesLogEntry.Pass
    GoTo Cleanup

TestFail:
    TestLogSolicitudOperation_EnrichesLogEntry.Fail "Error: " & Err.Description
    
Cleanup:
    Set serviceImpl = Nothing
    Set mockRepo = Nothing
    Set mockErrorHandler = Nothing
    Set service = Nothing
    Set logEntry = Nothing
    Set mockSolicitud = Nothing
End Function
