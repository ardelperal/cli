Attribute VB_Name = "TIMapeoRepository"
Option Compare Database
Option Explicit

' --- Constantes eliminadas - ahora se usa modTestUtils.GetWorkspacePath() ---

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE (ESTÁNDAR DE ORO)
' ============================================================================

Public Function TIMapeoRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIMapeoRepository (Estándar de Oro)"
    
    On Error GoTo CleanupSuite
    
    Call SuiteSetup
    suiteResult.AddResult TestGetMapeoPorTipoSuccess()
    suiteResult.AddResult TestGetMapeoPorTipoNotFound()
    
CleanupSuite:
    
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    
    Set TIMapeoRepositoryRunAll = suiteResult
End Function

' ============================================================================
' PROCEDIMIENTOS HELPER DE LA SUITE
' ============================================================================

Private Sub SuiteSetup()
    On Error GoTo ErrorHandler
    
    ' Insertar los datos de prueba maestros para la suite
    Dim db As DAO.Database
    Dim config As IConfig: Set config = modTestContext.GetTestConfig()
    Dim activePath As String: activePath = config.GetValue("CONDOR_DATA_PATH")
    Set db = DBEngine.OpenDatabase(activePath)
    
    ' BLINDAJE DE IDEMPOTENCIA: Limpiar datos de una ejecución anterior
    db.Execute "DELETE FROM tbMapeoCampos WHERE nombrePlantilla = 'PC' AND nombreCampoTabla = 'refContrato'", dbFailOnError
    
    db.Execute "INSERT INTO tbMapeoCampos (nombrePlantilla, nombreCampoTabla, nombreCampoWord) VALUES ('PC', 'refContrato', 'MARCADOR_CONTRATO')", dbFailOnError
    db.Close
    Set db = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "TIMapeoRepository.SuiteSetup", Err.Description
End Sub



' ============================================================================
' TESTS INDIVIDUALES (SE AÑADIRÁN EN LOS SIGUIENTES PROMPTS)
' ============================================================================

Private Function TestGetMapeoPorTipoSuccess() As CTestResult
    Set TestGetMapeoPorTipoSuccess = New CTestResult
    TestGetMapeoPorTipoSuccess.Initialize "GetMapeoPorTipo debe devolver un objeto EMapeo con datos"
    
    Dim repository As IMapeoRepository
    Dim mapeoResult As EMapeo
    
    On Error GoTo TestFail

    ' Arrange: Usar configuración centralizada
    Set repository = modRepositoryFactory.CreateMapeoRepository(modTestContext.GetTestConfig())
    
    ' Act
    Set mapeoResult = repository.GetMapeoPorTipo("PC")
    
    ' Assert
    modAssert.AssertNotNull mapeoResult, "El objeto EMapeo no debe ser nulo."
    modAssert.AssertEquals "PC", mapeoResult.NombrePlantilla, "El nombre de la plantilla no es el esperado."

    TestGetMapeoPorTipoSuccess.Pass
    GoTo Cleanup

TestFail:
    TestGetMapeoPorTipoSuccess.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set mapeoResult = Nothing
    Set repository = Nothing
End Function

Private Function TestGetMapeoPorTipoNotFound() As CTestResult
    Set TestGetMapeoPorTipoNotFound = New CTestResult
    TestGetMapeoPorTipoNotFound.Initialize "GetMapeoPorTipo debe devolver Nothing si no hay mapeo"
    
    Dim repository As IMapeoRepository
    Dim mapeoResult As EMapeo
    
    On Error GoTo TestFail

    ' Arrange: Usar configuración centralizada
    Set repository = modRepositoryFactory.CreateMapeoRepository(modTestContext.GetTestConfig())
    
    ' Act
    Set mapeoResult = repository.GetMapeoPorTipo("TIPO_INEXISTENTE")
    
    ' Assert
    modAssert.AssertIsNull mapeoResult, "El objeto EMapeo devuelto debería ser Nothing."

    TestGetMapeoPorTipoNotFound.Pass
    GoTo Cleanup

TestFail:
    TestGetMapeoPorTipoNotFound.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set mapeoResult = Nothing
    Set repository = Nothing
End Function
