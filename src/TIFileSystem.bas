Attribute VB_Name = "TIFileSystem"
Option Compare Database
Option Explicit

Private Const TEST_DIR As String = "fs_tests\"

' ============================================================================
' FUNCIÓN PRINCIPAL DE LA SUITE (ESTÁNDAR DE ORO)
' ============================================================================

Public Function TIFileSystemRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TIFileSystem (Estándar de Oro)"
    
    On Error GoTo CleanupSuite
    
    Call SuiteSetup
    suiteResult.AddResult TestCreateAndFolderExists()
    suiteResult.AddResult TestCreateAndDeleteFile()
    
CleanupSuite:
    
    If Err.Number <> 0 Then
        Dim errorTest As New CTestResult
        errorTest.Initialize "Suite_Execution_Failed"
        errorTest.Fail "La suite falló de forma catastrófica: " & Err.Description
        suiteResult.AddResult errorTest
    End If
    
    Set TIFileSystemRunAll = suiteResult
End Function

' ============================================================================
' PROCEDIMIENTOS HELPER DE LA SUITE
' ============================================================================

Private Sub SuiteSetup()
    ' Asegurarse de que el entorno está limpio antes de empezar
    
    
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    fs.CreateFolder modTestUtils.GetWorkspacePath() & TEST_DIR
    Set fs = Nothing
End Sub


' ============================================================================
' TESTS INDIVIDUALES
' ============================================================================

Private Function TestCreateAndFolderExists() As CTestResult
    Set TestCreateAndFolderExists = New CTestResult
    TestCreateAndFolderExists.Initialize "CreateFolder y FolderExists deben funcionar"
    
    Dim fs As IFileSystem
    On Error GoTo TestFail
    
    ' Arrange
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    ' Assert
    modAssert.AssertTrue fs.FolderExists(modTestUtils.GetWorkspacePath() & TEST_DIR), "La carpeta de prueba debería existir después del SuiteSetup."
    
    TestCreateAndFolderExists.Pass
    Exit Function
    
TestFail:
    TestCreateAndFolderExists.Fail Err.Description
End Function

Private Function TestCreateAndDeleteFile() As CTestResult
    Set TestCreateAndDeleteFile = New CTestResult
    TestCreateAndDeleteFile.Initialize "CopyFile, FileExists y DeleteFile deben funcionar"
    
    Dim fs As IFileSystem
    Dim testFilePath As String
    Dim fso As Object
    On Error GoTo TestFail
    
    ' Arrange
    Set fs = modFileSystemFactory.CreateFileSystem()
    testFilePath = modTestUtils.GetWorkspacePath() & TEST_DIR & "test_file.txt"
    
    ' Crear un fichero dummy para copiar
    Dim tempSourcePath As String
    tempSourcePath = modTestUtils.GetWorkspacePath() & TEST_DIR & "source.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateTextFile(tempSourcePath, True).Write "test"
    
    ' Act
    fs.CopyFile tempSourcePath, testFilePath
    
    ' Assert
    modAssert.AssertTrue fs.FileExists(testFilePath), "El archivo copiado debería existir."
    
    fs.DeleteFile testFilePath
    modAssert.AssertFalse fs.FileExists(testFilePath), "El archivo eliminado no debería existir."
    
    TestCreateAndDeleteFile.Pass
    GoTo Cleanup

TestFail:
    TestCreateAndDeleteFile.Fail Err.Description
    
Cleanup:
    Set fs = Nothing
    Set fso = Nothing
End Function
