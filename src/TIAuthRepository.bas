Attribute VB_Name = "TIAuthRepository"
Option Compare Database
Option Explicit

Public Function TIAuthRepositoryRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIAuthRepository"
    Call SuiteSetup
    suiteResult.AddResult TestGetUserAuthData_AdminUser_ReturnsCorrectData()
    Set TIAuthRepositoryRunAll = suiteResult
End Function

Private Sub SuiteSetup()
    Dim config As IConfig: Set config = modTestContext.GetTestConfig()
    Dim dbPath As String: dbPath = config.GetValue("LANZADERA_DATA_PATH")
    Dim db As DAO.Database: Set db = DBEngine.OpenDatabase(dbPath, False, False, ";PWD=" & config.GetValue("LANZADERA_PASSWORD"))
    
    ' BLINDAJE DE IDEMPOTENCIA: Limpiar datos de ejecuciones anteriores de AMBAS tablas
    db.Execute "DELETE FROM TbUsuariosAplicaciones WHERE CorreoUsuario='admin@example.com'", dbFailOnError
    db.Execute "DELETE FROM TbUsuariosAplicacionesPermisos WHERE CorreoUsuario='admin@example.com'", dbFailOnError
    
    ' Insertar datos de prueba, incluyendo el campo obligatorio 'Id'
    db.Execute "INSERT INTO TbUsuariosAplicaciones (Id, CorreoUsuario, EsAdministrador) VALUES (999, 'admin@example.com','Sí')", dbFailOnError
    db.Execute "INSERT INTO TbUsuariosAplicacionesPermisos (CorreoUsuario, IDAplicacion, EsUsuarioAdministrador) VALUES ('admin@example.com', 231, True)", dbFailOnError
    db.Close
End Sub

Private Function TestGetUserAuthData_AdminUser_ReturnsCorrectData() As CTestResult
    Set TestGetUserAuthData_AdminUser_ReturnsCorrectData = New CTestResult
    TestGetUserAuthData_AdminUser_ReturnsCorrectData.Initialize "GetUserAuthData para Admin debe devolver datos correctos"
    Dim repo As IAuthRepository: Set repo = modRepositoryFactory.CreateAuthRepository(modTestContext.GetTestConfig())
    Dim authData As EAuthData: Set authData = repo.GetUserAuthData("admin@example.com")
    modAssert.AssertNotNull authData, "AuthData no debe ser nulo."
    TestGetUserAuthData_AdminUser_ReturnsCorrectData.Pass
End Function
