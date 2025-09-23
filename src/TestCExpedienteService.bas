Attribute VB_Name = "TestCExpedienteService"
Option Compare Database
Option Explicit

'===============================================================================
' MÓDULO: TestCExpedienteService
' DESCRIPCIÓN: Pruebas unitarias para la clase CExpedienteService.
'              Verifica la correcta delegación de llamadas al repositorio.
' ESTÁNDAR: Oro
'===============================================================================

#If DEV_MODE Then


'===============================================================================
'                    FUNCIÓN PRINCIPAL DE LA SUITE DE PRUEBAS
'===============================================================================

Public Function TestCExpedienteServiceRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TestCExpedienteService"

    suiteResult.AddResult Test_ObtenerExpedientePorId_DelegatesCorrectly()
    
    Set TestCExpedienteServiceRunAll = suiteResult
End Function

'===============================================================================
'                             PRUEBAS INDIVIDUALES
'===============================================================================

Private Function Test_ObtenerExpedientePorId_DelegatesCorrectly() As CTestResult
    Set Test_ObtenerExpedientePorId_DelegatesCorrectly = New CTestResult
    Test_ObtenerExpedientePorId_DelegatesCorrectly.Initialize "Debe delegar la llamada a ObtenerExpedientePorId del repositorio con el ID correcto"
    
    ' --- Declaración de Variables ---
    Dim service As IExpedienteService
    Dim serviceImpl As CExpedienteService
    Dim mockRepo As CMockExpedienteRepository
    Dim mockLogger As CMockOperationLogger
    Dim mockErrorHandler As CMockErrorHandlerService
    Dim expectedExpediente As EExpediente
    Dim actualExpediente As EExpediente
    Dim testId As Long
    
    On Error GoTo TestFail
    
    ' --- ARRANGE ---
    ' 1. Crear dependencias mockeadas
    Set mockRepo = New CMockExpedienteRepository
    Set mockLogger = New CMockOperationLogger
    Set mockErrorHandler = New CMockErrorHandlerService
    
    ' 2. Configurar el mock del repositorio
    testId = 123
    Set expectedExpediente = New EExpediente
    expectedExpediente.idExpediente = testId
    expectedExpediente.Nemotecnico = "NEMO-TEST"
    mockRepo.ConfigureObtenerExpedientePorId expectedExpediente
    
    ' 3. Crear e inicializar el objeto bajo prueba (Service)
    Set serviceImpl = New CExpedienteService
    serviceImpl.Initialize mockRepo, mockLogger, mockErrorHandler
    Set service = serviceImpl
    
    ' --- ACT ---
    ' Llamar al método del servicio que queremos probar
    Set actualExpediente = service.ObtenerExpedientePorId(testId)
    
    ' --- ASSERT ---
    ' 1. Verificar que el mock fue llamado
    modAssert.AssertTrue mockRepo.ObtenerExpedientePorId_WasCalled, "El método ObtenerExpedientePorId del repositorio debería haber sido llamado."
    
    ' 2. Verificar que el mock fue llamado CON LOS PARÁMETROS CORRECTOS
    modAssert.AssertEquals testId, mockRepo.ObtenerExpedientePorId_LastId, "El servicio no pasó el ID correcto al repositorio."
    
    ' 3. Verificar que el servicio devolvió el objeto que el mock le proporcionó
    modAssert.AssertEquals expectedExpediente.Nemotecnico, actualExpediente.Nemotecnico, "El expediente devuelto no es el esperado."
    
    Test_ObtenerExpedientePorId_DelegatesCorrectly.Pass
    GoTo Cleanup

TestFail:
    Test_ObtenerExpedientePorId_DelegatesCorrectly.Fail "Error inesperado #" & Err.Number & ": " & Err.Description
    
Cleanup:
    ' Limpieza de recursos
    Set service = Nothing
    Set serviceImpl = Nothing
    Set mockRepo = Nothing
    Set mockLogger = Nothing
    Set mockErrorHandler = Nothing
    Set expectedExpediente = Nothing
    Set actualExpediente = Nothing
End Function

#End If
