Attribute VB_Name = "TIOperationRepository"
Option Compare Database
Option Explicit

' ============================================================================
' MÓDULO: TIOperationRepository
' DESCRIPCIÓN: Suite de pruebas de integración para COperationRepository
' ARQUITECTURA: Patrón de Oro (Setup a Nivel de Suite + Transacciones)
' ============================================================================

' --- Constantes eliminadas - ahora se usa modTestUtils.GetWorkspacePath() ---

' ============================================================================
' FUNCIÓN PRINCIPAL DE EJECUCIÓN
' ============================================================================

Public Function TIOperationRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIOperationRepository (Estándar de Oro)"
    
    On Error GoTo CleanupSuite
    
    Call SuiteSetup
    suiteResult.AddResult TestSaveLog_Success()
    
CleanupSuite:
    
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    
    Set TIOperationRepositoryRunAll = suiteResult
End Function

' ============================================================================
' PROCEDIMIENTOS HELPER DE LA SUITE
' ============================================================================

Private Sub SuiteSetup()
    On Error GoTo ErrorHandler

    ' Conectar a la BD y limpiar los datos de prueba específicos de esta suite
    Dim config As IConfig: Set config = modTestContext.GetTestConfig()
    Dim dbPath As String: dbPath = config.GetValue("CONDOR_DATA_PATH")
    Dim db As DAO.Database: Set db = DBEngine.OpenDatabase(dbPath, False, False)
    
    ' BLINDAJE DE IDEMPOTENCIA: Eliminar el registro de prueba por su tipo único
    db.Execute "DELETE FROM tbOperacionesLog WHERE tipoOperacion = 'TEST_OP'", dbFailOnError
    
    db.Close
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "TIOperationRepository.SuiteSetup", Err.Description
End Sub



' ============================================================================
' PRUEBAS DE INTEGRACIÓN
' ============================================================================

Private Function TestSaveLog_Success() As CTestResult
    Set TestSaveLog_Success = New CTestResult
    TestSaveLog_Success.Initialize "SaveLog debe guardar correctamente un EOperationLog"
    
    Dim repository As IOperationRepository
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim config As IConfig
    Dim logEntry As EOperationLog
    
    On Error GoTo TestFail

    ' ARRANGE:
    Set config = modTestContext.GetTestConfig()
    Set repository = modRepositoryFactory.CreateOperationRepository(config)
    
    ' Crear el objeto de entidad a guardar
    Set logEntry = New EOperationLog
    logEntry.Initialize Now, "test_user", "TEST_OP", "Solicitud", 123, "Descripción de prueba.", "SUCCESS", "Detalles de prueba."

    ' ACT: Ejecutar la operación a probar
    repository.SaveLog logEntry
    
    ' ASSERT: Verificar directamente en la BD
    Set db = DBEngine.OpenDatabase(config.GetCondorDataPath(), False, False)
    Set rs = db.OpenRecordset("SELECT * FROM tbOperacionesLog WHERE tipoOperacion = 'TEST_OP'")
    
    modAssert.AssertFalse rs.EOF, "Se debería haber insertado un registro de log."
    modAssert.AssertEquals 123, rs!IdEntidad.value, "El ID de entidad no coincide."
    modAssert.AssertEquals "Solicitud", rs!entidad.value, "La entidad no coincide."
    modAssert.AssertEquals "Descripción de prueba.", rs!descripcion.value, "La descripción no coincide."
    modAssert.AssertEquals "SUCCESS", rs!resultado.value, "El resultado no coincide."
    modAssert.AssertEquals "Detalles de prueba.", rs!detalles.value, "Los detalles no coinciden."

    TestSaveLog_Success.Pass
    GoTo Cleanup

TestFail:
    TestSaveLog_Success.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing
    Set db = Nothing
    Set repository = Nothing
    Set config = Nothing
    Set logEntry = Nothing
End Function
