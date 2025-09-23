Attribute VB_Name = "TISolicitudRepository"
Option Compare Database
Option Explicit

' --- Constantes eliminadas - ahora se usa modTestUtils.GetWorkspacePath() ---

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE (ESTÁNDAR DE ORO)
' ============================================================================

Public Function TISolicitudRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TISolicitudRepository (Estándar de Oro)"
    
    On Error GoTo CleanupSuite
    
    Call SuiteSetup
    suiteResult.AddResult TestSaveAndRetrieveSolicitud()
    
CleanupSuite:
    
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    
    Set TISolicitudRepositoryRunAll = suiteResult
End Function

' ============================================================================
' PROCEDIMIENTOS HELPER DE LA SUITE
' ============================================================================

Private Sub SuiteSetup()
    On Error GoTo ErrorHandler
    Dim config As IConfig: Set config = modTestContext.GetTestConfig()
    Dim activePath As String: activePath = config.GetValue("CONDOR_DATA_PATH")
    Dim db As DAO.Database: Set db = DBEngine.OpenDatabase(activePath, False, False)
    
    ' BLINDAJE DE IDEMPOTENCIA:
    db.Execute "DELETE FROM tbSolicitudes WHERE codigoSolicitud = 'TEST-SAVE-001'", dbFailOnError
    
    db.Close
    Set db = Nothing
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, "TISolicitudRepository.SuiteSetup", Err.Description
End Sub



' ============================================================================
' PRUEBAS DE INTEGRACIÓN
' ============================================================================

Private Function TestSaveAndRetrieveSolicitud() As CTestResult
    Set TestSaveAndRetrieveSolicitud = New CTestResult
    TestSaveAndRetrieveSolicitud.Initialize "Debe guardar y recuperar una solicitud correctamente"
    
    Dim repo As ISolicitudRepository
    Dim retrievedSolicitud As ESolicitud
    
    On Error GoTo TestFail

    ' Arrange
    Set repo = modRepositoryFactory.CreateSolicitudRepository(modTestContext.GetTestConfig())
    
    Dim nuevaSolicitud As New ESolicitud
    nuevaSolicitud.idExpediente = 999
    nuevaSolicitud.tipoSolicitud = "TIPO_TEST"
    nuevaSolicitud.codigoSolicitud = "TEST-SAVE-001"
    nuevaSolicitud.idEstadoInterno = 1 ' Borrador
    nuevaSolicitud.usuarioCreacion = "itest_user"
    
    ' Act
    Dim newId As Long
    newId = repo.SaveSolicitud(nuevaSolicitud)
    Set retrievedSolicitud = repo.ObtenerSolicitudPorId(newId)
    
    ' Assert
    modAssert.AssertTrue newId > 0, "El ID devuelto debe ser positivo."
    modAssert.AssertNotNull retrievedSolicitud, "La solicitud recuperada no debe ser nula."
    modAssert.AssertEquals "TEST-SAVE-001", retrievedSolicitud.codigoSolicitud, "El código no coincide."

    TestSaveAndRetrieveSolicitud.Pass
    GoTo Cleanup

TestFail:
    TestSaveAndRetrieveSolicitud.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set retrievedSolicitud = Nothing
    Set repo = Nothing
End Function
