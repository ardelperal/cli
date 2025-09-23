Attribute VB_Name = "TestErrorHandlerService"
Option Compare Database
Option Explicit


' =====================================================
' TestErrorHandlerService.bas
' Módulo de pruebas para CErrorHandlerService
' =====================================================

Public Function TestErrorHandlerServiceRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TestErrorHandlerService"
    
    suiteResult.AddResult TestLogError_WritesToFile_Success()
    
    Set TestErrorHandlerServiceRunAll = suiteResult
End Function

Private Function TestLogError_WritesToFile_Success() As CTestResult
    Set TestLogError_WritesToFile_Success = New CTestResult
    TestLogError_WritesToFile_Success.Initialize "LogError debe escribir en el fichero de log correctamente"
    
    Dim serviceImpl As CErrorHandlerService
    Dim mockConfig As CMockConfig
    Dim mockFileSystem As CMockFileSystem
    Dim service As IErrorHandlerService
    
    On Error GoTo TestFail
    
    ' Arrange
    Set mockConfig = New CMockConfig
    Set mockFileSystem = New CMockFileSystem
    mockConfig.SetSetting "LOG_FILE_PATH", "C:\fake\log.txt"
    
    Set serviceImpl = New CErrorHandlerService
    serviceImpl.Initialize mockConfig, mockFileSystem
    Set service = serviceImpl
    
    ' Act
    service.LogError 123, "Test Error", "TestModule"
    
    ' Assert
    modAssert.AssertTrue mockFileSystem.WriteLineToFile_WasCalled, "WriteLineToFile debería haber sido llamado."
    modAssert.AssertEquals "C:\fake\log.txt", mockFileSystem.WriteLineToFile_LastPath, "Se llamó a WriteLineToFile con la ruta incorrecta."
    modAssert.AssertTrue InStr(mockFileSystem.WriteLineToFile_LastLine, "Test Error") > 0, "El mensaje de error no se escribió correctamente en la línea."
    
    TestLogError_WritesToFile_Success.Pass
    GoTo Cleanup
    
TestFail:
    TestLogError_WritesToFile_Success.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set serviceImpl = Nothing
    Set mockConfig = Nothing
    Set mockFileSystem = Nothing
    Set service = Nothing
End Function
