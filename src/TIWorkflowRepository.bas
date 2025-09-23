Attribute VB_Name = "TIWorkflowRepository"
Option Compare Database
Option Explicit

Public Function TIWorkflowRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIWorkflowRepository (Estándar de Oro)"
    
    On Error GoTo CleanupSuite
    
    Call SuiteSetup
    suiteResult.AddResult TestIsValidTransition_TrueForValidPath()
    suiteResult.AddResult TestIsValidTransition_FalseForInvalidPath()
    suiteResult.AddResult TestGetNextStates_ReturnsCorrectStates()
    
CleanupSuite:
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    
    Set TIWorkflowRepositoryRunAll = suiteResult
End Function

Private Sub SuiteSetup()
    On Error GoTo ErrorHandler
    Dim db As DAO.Database
    Dim config As IConfig: Set config = modTestContext.GetTestConfig()
    Dim activePath As String: activePath = config.GetValue("CONDOR_DATA_PATH")
    Set db = DBEngine.OpenDatabase(activePath)
    db.Execute "DELETE FROM tbTransiciones", dbFailOnError
    db.Execute "DELETE FROM tbEstados", dbFailOnError
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado) VALUES (1, 'Registrado');", dbFailOnError
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado) VALUES (2, 'Desarrollo');", dbFailOnError
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado) VALUES (3, 'Modificacion');", dbFailOnError
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado) VALUES (4, 'Validacion');", dbFailOnError
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado) VALUES (5, 'Revision');", dbFailOnError
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado) VALUES (6, 'Formalizacion');", dbFailOnError
    db.Execute "INSERT INTO tbEstados (idEstado, nombreEstado) VALUES (7, 'Aprobada');", dbFailOnError
     db.Execute "INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (1, 1, 2, 'Calidad');", dbFailOnError
    db.Execute "INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (2, 2, 3, 'Tecnico');", dbFailOnError
    db.Execute "INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (3, 3, 4, 'Calidad');", dbFailOnError
    db.Execute "INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (4, 4, 5, 'Calidad');", dbFailOnError
    db.Execute "INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (5, 4, 5, 'Tecnico');", dbFailOnError
    db.Execute "INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (6, 5, 6, 'Calidad');", dbFailOnError
    db.Execute "INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (7, 5, 3, 'Calidad');", dbFailOnError
    db.Execute "INSERT INTO tbTransiciones (idTransicion, idEstadoOrigen, idEstadoDestino, rolRequerido) VALUES (8, 6, 7, 'Calidad');", dbFailOnError
    db.Close
    Set db = Nothing
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "TIWorkflowRepository.SuiteSetup", Err.Description
End Sub

Private Function TestIsValidTransition_TrueForValidPath() As CTestResult
    Set TestIsValidTransition_TrueForValidPath = New CTestResult
    TestIsValidTransition_TrueForValidPath.Initialize "IsValidTransition debe devolver True para transiciones válidas del nuevo flujo"
    Dim repo As IWorkflowRepository
    On Error GoTo TestFail
    Set repo = modRepositoryFactory.CreateWorkflowRepository(modTestContext.GetTestConfig())
    
    modAssert.AssertTrue repo.IsValidTransition("", "Registrado", "Desarrollo", "Calidad"), "Calidad debe poder pasar de Registrado a Desarrollo."
    modAssert.AssertTrue repo.IsValidTransition("", "Desarrollo", "Modificacion", "Tecnico"), "Tecnico debe poder pasar de Desarrollo a Modificacion."
    modAssert.AssertTrue repo.IsValidTransition("", "Revision", "Formalizacion", "Calidad"), "Calidad debe poder pasar de Revision a Formalizacion."
    
    TestIsValidTransition_TrueForValidPath.Pass
    GoTo Cleanup
TestFail:
    TestIsValidTransition_TrueForValidPath.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set repo = Nothing
End Function

Private Function TestIsValidTransition_FalseForInvalidPath() As CTestResult
    Set TestIsValidTransition_FalseForInvalidPath = New CTestResult
    TestIsValidTransition_FalseForInvalidPath.Initialize "IsValidTransition debe devolver False para transiciones inválidas"
    Dim repo As IWorkflowRepository
    On Error GoTo TestFail
    Set repo = modRepositoryFactory.CreateWorkflowRepository(modTestContext.GetTestConfig())

    modAssert.AssertFalse repo.IsValidTransition("", "Registrado", "Desarrollo", "Tecnico"), "Un Tecnico NO debe poder pasar de Registrado a Desarrollo."
    modAssert.AssertFalse repo.IsValidTransition("", "Registrado", "Aprobada", "Calidad"), "Nadie debe poder saltar directamente a Aprobada."
    
    TestIsValidTransition_FalseForInvalidPath.Pass
    GoTo Cleanup
TestFail:
    TestIsValidTransition_FalseForInvalidPath.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set repo = Nothing
End Function

Private Function TestGetNextStates_ReturnsCorrectStates() As CTestResult
    Set TestGetNextStates_ReturnsCorrectStates = New CTestResult
    TestGetNextStates_ReturnsCorrectStates.Initialize "GetNextStates debe devolver los estados correctos para un rol"
    Dim repo As IWorkflowRepository, nextStates As Scripting.Dictionary
    On Error GoTo TestFail
    Set repo = modRepositoryFactory.CreateWorkflowRepository(modTestContext.GetTestConfig())
    
    ' Act: Para un Técnico en estado 'Validacion' (ID 4)
    Set nextStates = repo.GetNextStates(4, "Tecnico")
    
    ' Assert: Solo debe poder moverlo a 'Revision' (ID 5)
    modAssert.AssertEquals 1, nextStates.Count, "Un Tecnico en Validacion solo debe tener un estado siguiente."
    modAssert.AssertTrue nextStates.Exists(5), "El estado siguiente para un Tecnico en Validacion debe ser 'Revision' (ID 5)."
    
    TestGetNextStates_ReturnsCorrectStates.Pass
    GoTo Cleanup
TestFail:
    TestGetNextStates_ReturnsCorrectStates.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set nextStates = Nothing
    Set repo = Nothing
End Function
