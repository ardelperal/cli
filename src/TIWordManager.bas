Attribute VB_Name = "TIWordManager"
Option Compare Database
Option Explicit



' ============================================================================
' REQUISITO DE COMPILACIÓN CRÍTICO:
' "Microsoft Word XX.X Object Library" debe estar referenciada.
' ============================================================================

Private Const TEST_FOLDER_REL As String = "word_manager_tests\"

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE (ESTÁNDAR DE ORO)
' ============================================================================

Public Function TIWordManagerRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIWordManager (Estándar de Oro)"
    
    On Error GoTo CleanupSuite
    
    Call SuiteSetup
    suiteResult.AddResult Test_CicloCompleto_Success()
    suiteResult.AddResult Test_AbrirFicheroInexistente_DevuelveFalse()
    
CleanupSuite:
    Call SuiteTeardown
    
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    
    Set TIWordManagerRunAll = suiteResult
End Function

' ============================================================================
' PROCEDIMIENTOS HELPER DE LA SUITE
' ============================================================================

Private Sub SuiteSetup()
        On Error GoTo ErrorHandler
        
        ' 1. OBTENER CONFIGURACIÓN EXCLUSIVAMENTE DESDE EL CONTEXTO DE PRUEBAS (Lección 37)
        Dim config As IConfig: Set config = modTestContext.GetTestConfig()
        Dim fs As IFileSystem: Set fs = modFileSystemFactory.CreateFileSystem(config)
        
        ' 2. OBTENER RUTAS Y NOMBRES DE LA PLANTILLA REAL DE PRODUCCIÓN DESDE LA CONFIGURACIÓN
        Dim productionTemplatesPath As String: productionTemplatesPath = config.GetValue("PRODUCTION_TEMPLATES_PATH")
        Dim templateFilename As String: templateFilename = config.GetValue("TEMPLATE_CDCA_FILENAME")
        
        ' 3. CONSTRUIR RUTAS DE ORIGEN Y DESTINO
        Dim sourceTemplatePath As String: sourceTemplatePath = modTestUtils.JoinPath(productionTemplatesPath, templateFilename)
        Dim destinationFolder As String: destinationFolder = modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), TEST_FOLDER_REL)
        Dim destinationFilePath As String: destinationFilePath = modTestUtils.JoinPath(destinationFolder, templateFilename)
        
        ' 4. APROVISIONAR EL ENTORNO: LIMPIAR Y COPIAR
        ' Asegurarse de que la carpeta de destino existe y está vacía
        If fs.FolderExists(destinationFolder) Then fs.DeleteFolderRecursive destinationFolder
        fs.CreateFolder destinationFolder
        
        ' Verificar que la plantilla de producción (el fixture) existe antes de copiarla
        If Not fs.FileExists(sourceTemplatePath) Then
            Err.Raise vbObjectError + 5501, "SuiteSetup", "La plantilla de origen maestra no se encontró en: " & sourceTemplatePath
        End If
        
        ' Copiar la plantilla de producción al workspace de pruebas
        fs.CopyFile sourceTemplatePath, destinationFilePath
        
        Exit Sub

ErrorHandler:
        ' Si el setup falla, la suite entera debe fallar con un mensaje claro.
        Err.Raise vbObjectError + 5500, "SuiteSetup", "Fallo catastrófico al aprovisionar el entorno para TIWordManager: " & Err.Description
    End Sub

Private Sub SuiteTeardown()
    On Error Resume Next ' Ignorar errores durante la limpieza
    
    ' 1. Limpiar carpeta de pruebas
    Dim fs As IFileSystem: Set fs = modFileSystemFactory.CreateFileSystem()
    Dim testFolder As String: testFolder = modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), TEST_FOLDER_REL)
    fs.DeleteFolderRecursive testFolder
    Set fs = Nothing
    
    ' 2. Forzar cierre de todas las instancias de Word que puedan haber quedado abiertas
    Dim wordApp As Object
    Set wordApp = Nothing
    
    ' Intentar cerrar Word mediante automation si está disponible
    On Error Resume Next
    Set wordApp = GetObject(, "Word.Application")
    If Not wordApp Is Nothing Then
        wordApp.Quit SaveChanges:=False
        Set wordApp = Nothing
    End If
    On Error GoTo 0
End Sub



' ============================================================================
' TESTS INDIVIDUALES (SE AÑADIRÁN EN LOS SIGUIENTES PROMPTS)
' ============================================================================

Private Function Test_CicloCompleto_Success() As CTestResult
    Set Test_CicloCompleto_Success = New CTestResult
    Test_CicloCompleto_Success.Initialize "Ciclo completo debe usar plantilla real de /recursos/"
    
    Dim writerWordManager As IWordManager
    Dim readerWordManager As IWordManager
    Dim fs As IFileSystem
    Dim config As IConfig
    
    On Error GoTo TestFail
    
    ' --- Arrange ---
    ' 1. Obtener configuración.
    Set config = modTestContext.GetTestConfig()
    
    ' 2. Construir la ruta al documento de prueba que SuiteSetup ya ha aprovisionado.
    '    SuiteSetup ya ha copiado 'CD_CA.docx' a la carpeta 'word_manager_tests\'.
    Dim templateFilename As String: templateFilename = config.GetValue("TEMPLATE_CDCA_FILENAME")
    Dim workspacePath As String: workspacePath = modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), TEST_FOLDER_REL)
    Dim testDocPath As String: testDocPath = modTestUtils.JoinPath(workspacePath, templateFilename)
    
    ' --- Act (Escritura) ---
    Set writerWordManager = modWordManagerFactory.CreateWordManager()
    writerWordManager.AbrirDocumento testDocPath
    writerWordManager.SetBookmarkText "Parte0_1", "TEST-REF-SUMINISTRADOR"
    writerWordManager.GuardarDocumento
    writerWordManager.Dispose
    Set writerWordManager = Nothing
    
    ' --- Act (Lectura) ---
    Set readerWordManager = modWordManagerFactory.CreateWordManager()
    readerWordManager.AbrirDocumento testDocPath
    Dim contenidoLeido As String
    contenidoLeido = readerWordManager.GetBookmarkText("Parte0_1")
    
    ' --- Assert ---
    modAssert.AssertEquals "TEST-REF-SUMINISTRADOR", contenidoLeido, "El contenido leído del bookmark no es el esperado."
    Test_CicloCompleto_Success.Pass
    GoTo Cleanup

TestFail:
    Test_CicloCompleto_Success.Fail "Error: " & Err.Description
    
Cleanup:
    If Not writerWordManager Is Nothing Then writerWordManager.Dispose
    If Not readerWordManager Is Nothing Then readerWordManager.Dispose
    Set fs = Nothing
    Set config = Nothing
    Set writerWordManager = Nothing
    Set readerWordManager = Nothing
End Function

Private Function Test_AbrirFicheroInexistente_DevuelveFalse() As CTestResult
    Set Test_AbrirFicheroInexistente_DevuelveFalse = New CTestResult
    Test_AbrirFicheroInexistente_DevuelveFalse.Initialize "Abrir un fichero inexistente debe devolver False y no lanzar error"
    
    Dim wordManager As IWordManager
    On Error GoTo TestFail ' Si hay algún error, el test falla
    
    Set wordManager = modWordManagerFactory.CreateWordManager()
    Dim inexistentePath As String: inexistentePath = modTestUtils.JoinPath(modTestUtils.GetWorkspacePath() & "word_manager_tests\", "no_existe.docx")
    
    ' Act
    Dim result As Boolean
    result = wordManager.AbrirDocumento(inexistentePath)
    
    ' Assert
    modAssert.AssertFalse result, "AbrirDocumento debería haber devuelto False."
    
    Test_AbrirFicheroInexistente_DevuelveFalse.Pass
    GoTo Cleanup

TestFail:
    Test_AbrirFicheroInexistente_DevuelveFalse.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    If Not wordManager Is Nothing Then wordManager.Dispose
    Set wordManager = Nothing
End Function
