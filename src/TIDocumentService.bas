Attribute VB_Name = "TIDocumentService"
Option Compare Database
Option Explicit

' =====================================================
' Módulo: TIDocumentService
' Descripción: Pruebas de integración para CDocumentService
' Versión: 3.1 (Consolidada y robustecida)
' =====================================================

Private Sub SuiteSetup()
    On Error GoTo ErrorHandler
    
    Dim fs As IFileSystem
    Dim config As IConfig
    Dim db As DAO.Database
    
    ' 1. OBTENER CONFIGURACIÓN Y SISTEMA DE FICHEROS
    Set config = modTestContext.GetTestConfig()
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    ' 2. PREPARAR SISTEMA DE ARCHIVOS Y APROVISIONAR PLANTILLA
    Dim testTemplatesPath As String: testTemplatesPath = config.GetValue("TEMPLATES_PATH")
    fs.CreateFolder testTemplatesPath ' Crear la carpeta de plantillas en el workspace
    
    Dim productionTemplatePath As String: productionTemplatePath = config.GetValue("PRODUCTION_TEMPLATES_PATH") & config.GetValue("TEMPLATE_PC_FILENAME")
    Dim testTemplatePath As String: testTemplatePath = testTemplatesPath & config.GetValue("TEMPLATE_PC_FILENAME")
    
    fs.CopyFile productionTemplatePath, testTemplatePath ' Copiar la plantilla para el test
    
    ' 3. PREPARAR BASE DE DATOS (SEMBRADO IDEMPOTENTE)
    Dim dbPath As String: dbPath = config.GetCondorDataPath()
    
    If Not fs.FileExists(dbPath) Then
        Err.Raise vbObjectError + 1001, "TIDocumentService.SuiteSetup", "La base de datos de prueba no fue encontrada en la ruta: " & dbPath
    End If
    
    Set db = DBEngine.OpenDatabase(dbPath, False, False)
    
    db.Execute "DELETE FROM tbSolicitudes WHERE idSolicitud = 1", dbFailOnError
    db.Execute "DELETE FROM tbMapeoCampos WHERE nombrePlantilla = 'PC'", dbFailOnError
    
    db.Execute "INSERT INTO tbSolicitudes (idSolicitud, idExpediente, tipoSolicitud, codigoSolicitud, idEstadoInterno, usuarioCreacion, fechaCreacion) " & _
               "VALUES (1, 999, 'PC', 'COD-001', 1, 'test_user', Now())", dbFailOnError
               
    db.Execute "INSERT INTO tbMapeoCampos (nombrePlantilla, nombreCampoTabla, nombreCampoWord) " & _
               "VALUES ('PC', 'codigoSolicitud', 'BookmarkCodigo')", dbFailOnError
    db.Execute "INSERT INTO tbMapeoCampos (nombrePlantilla, nombreCampoTabla, nombreCampoWord) " & _
               "VALUES ('PC', 'usuarioCreacion', 'BookmarkUsuario')", dbFailOnError
               
    db.Close
    
    Exit Sub
    
ErrorHandler:
    If Not db Is Nothing Then db.Close
    Set db = Nothing
    Set fs = Nothing
    Set config = Nothing
    Err.Raise Err.Number, "TIDocumentService.SuiteSetup", Err.Description
End Sub

Public Function TIDocumentServiceRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIDocumentService"
    
    On Error GoTo CleanupSuite

    Call SuiteSetup
    suiteResult.AddResult TestGenerarDocumentoSuccess()
    
CleanupSuite:
    Call SuiteTeardown
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Error"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    
    Set TIDocumentServiceRunAll = suiteResult
End Function

Private Function TestGenerarDocumentoSuccess() As CTestResult
    Set TestGenerarDocumentoSuccess = New CTestResult
    TestGenerarDocumentoSuccess.Initialize "GenerarDocumento debe crear un archivo Word con datos reales de la BD"

    Dim documentService As IDocumentService
    Dim fs As IFileSystem
    Dim config As IConfig
    
    On Error GoTo TestFail

    Set config = modTestContext.GetTestConfig()
    Set fs = modFileSystemFactory.CreateFileSystem(config)
    Set documentService = modDocumentServiceFactory.CreateDocumentService(config)
    
    Dim rutaGenerada As String
    rutaGenerada = documentService.GenerarDocumento(1)

    modAssert.AssertNotEquals "", rutaGenerada, "La ruta del documento generado no debe estar vacía."
    modAssert.AssertTrue fs.FileExists(rutaGenerada), "El archivo generado debe existir en el disco."
    
    TestGenerarDocumentoSuccess.Pass
    GoTo Cleanup

TestFail:
    TestGenerarDocumentoSuccess.Fail "Error: " & Err.Description
    
Cleanup:
    Set documentService = Nothing
    Set fs = Nothing
    Set config = Nothing
End Function

Private Sub SuiteTeardown()
    On Error Resume Next
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    fs.DeleteFolderRecursive modTestUtils.GetWorkspacePath() & "doc_service_test\"
    Set fs = Nothing
End Sub


