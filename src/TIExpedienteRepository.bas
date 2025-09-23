Attribute VB_Name = "TIExpedienteRepository"

Option Compare Database
Option Explicit

Public Function TIExpedienteRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIExpedienteRepository"
    Call SuiteSetup
    suiteResult.AddResult Test_ObtenerExpedientePorId_Exitoso()
    Set TIExpedienteRepositoryRunAll = suiteResult
End Function

Private Sub SuiteSetup()
    Dim config As IConfig: Set config = modTestContext.GetTestConfig()
    Dim dbPath As String: dbPath = config.GetValue("EXPEDIENTES_DATA_PATH")
    Dim db As DAO.Database: Set db = DBEngine.OpenDatabase(dbPath, False, False, ";PWD=" & config.GetValue("EXPEDIENTES_PASSWORD"))
    db.Execute "DELETE FROM TbExpedientes WHERE IDExpediente = 1", dbFailOnError
    db.Execute "INSERT INTO TbExpedientes (IDExpediente, Nemotecnico) VALUES (1, 'NEMO-001')", dbFailOnError
    db.Close
End Sub

Private Function Test_ObtenerExpedientePorId_Exitoso() As CTestResult
    Set Test_ObtenerExpedientePorId_Exitoso = New CTestResult
    Test_ObtenerExpedientePorId_Exitoso.Initialize "ObtenerExpedientePorId debe devolver un EExpediente poblado"
    Dim repo As IExpedienteRepository: Set repo = modRepositoryFactory.CreateExpedienteRepository(modTestContext.GetTestConfig())
    Dim result As EExpediente: Set result = repo.ObtenerExpedientePorId(1)
    modAssert.AssertNotNull result, "El expediente no debe ser Nulo."
    modAssert.AssertEquals "NEMO-001", result.Nemotecnico, "El nemotécnico no coincide."
    Test_ObtenerExpedientePorId_Exitoso.Pass
End Function
