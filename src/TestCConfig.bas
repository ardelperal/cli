Attribute VB_Name = "TestCConfig"
Option Compare Database
Option Explicit


' ============================================================================
' MÓDULO DE PRUEBAS UNITARIAS PARA CConfig
' Arquitectura: Pruebas Aisladas contra la implementación REAL de CConfig,
' utilizando LoadFromDictionary para inyectar configuración en memoria.
' ============================================================================

Public Function TestCConfigRunAll() As CTestSuiteResult
    Dim suiteResult As New CTestSuiteResult
    suiteResult.Initialize "TestCConfig - Pruebas Unitarias CConfig (Reconstruido)"
    
    suiteResult.AddResult TestGetCondorDataPath_ReturnsCorrectValue()
    suiteResult.AddResult TestHasKey_ReturnsTrueForExistingKey()
    suiteResult.AddResult TestHasKey_ReturnsFalseForNonExistingKey()
    suiteResult.AddResult TestGetValue_ReturnsEmptyForNonExistingKey()
    
    Set TestCConfigRunAll = suiteResult
End Function

Private Function TestGetCondorDataPath_ReturnsCorrectValue() As CTestResult
    Set TestGetCondorDataPath_ReturnsCorrectValue = New CTestResult
    TestGetCondorDataPath_ReturnsCorrectValue.Initialize "GetCondorDataPath debe devolver el valor correcto cargado"
    
    Dim config As CConfig
    Dim testSettings As Object
    On Error GoTo TestFail
    
    ' Arrange
    Set config = New CConfig
    Set testSettings = New Scripting.Dictionary
    testSettings.CompareMode = TextCompare
    testSettings.Add "CONDOR_DATA_PATH", "C:\Ruta\De\Prueba.accdb"
    config.LoadFromDictionary testSettings
    
    ' Act
    Dim result As String
    result = config.GetCondorDataPath()
    
    ' Assert
    modAssert.AssertEquals "C:\Ruta\De\Prueba.accdb", result, "GetCondorDataPath no devolvió el valor inyectado."
    
    TestGetCondorDataPath_ReturnsCorrectValue.Pass
    GoTo Cleanup
    
TestFail:
    TestGetCondorDataPath_ReturnsCorrectValue.Fail "Error inesperado: " & Err.Description
    
Cleanup:
    Set config = Nothing
    Set testSettings = Nothing
End Function



Private Function TestHasKey_ReturnsTrueForExistingKey() As CTestResult
    Set TestHasKey_ReturnsTrueForExistingKey = New CTestResult
    TestHasKey_ReturnsTrueForExistingKey.Initialize "HasKey debe devolver True para una clave existente"

    Dim config As CConfig
    Dim testSettings As Object
    On Error GoTo TestFail

    ' Arrange
    Set config = New CConfig
    Set testSettings = New Scripting.Dictionary
    testSettings.CompareMode = TextCompare
    testSettings.Add "EXISTING_KEY", "value"
    config.LoadFromDictionary testSettings

    ' Act
        Dim result As Boolean
    result = (config.GetValue("EXISTING_KEY") <> "")
    ' Assert
    modAssert.AssertTrue result, "La comprobación GetValue <> '' debería haber devuelto True para una clave existente."
    

    TestHasKey_ReturnsTrueForExistingKey.Pass
    GoTo Cleanup

TestFail:
    TestHasKey_ReturnsTrueForExistingKey.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set config = Nothing
    Set testSettings = Nothing
End Function

Private Function TestHasKey_ReturnsFalseForNonExistingKey() As CTestResult
    Set TestHasKey_ReturnsFalseForNonExistingKey = New CTestResult
    TestHasKey_ReturnsFalseForNonExistingKey.Initialize "HasKey (via GetValue<>'') debe devolver False para una clave inexistente"
    
    Dim config As CConfig
    On Error GoTo TestFail

    ' Arrange: Usamos un objeto de configuración vacío
    Set config = New CConfig

    ' Act: Usamos el patrón GetValue <> "" para verificar la existencia de la clave
    Dim result As Boolean
    result = (config.GetValue("NON_EXISTING_KEY") <> "")

    ' Assert: El resultado de la comparación debe ser False,
    ' ya que GetValue devuelve "" y la expresión ("" <> "") es False.
    modAssert.AssertFalse result, "La comprobación GetValue <> '' debería haber devuelto False."

    TestHasKey_ReturnsFalseForNonExistingKey.Pass
    GoTo Cleanup

TestFail:
    TestHasKey_ReturnsFalseForNonExistingKey.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set config = Nothing
End Function

Private Function TestGetValue_ReturnsEmptyForNonExistingKey() As CTestResult
    Set TestGetValue_ReturnsEmptyForNonExistingKey = New CTestResult
    TestGetValue_ReturnsEmptyForNonExistingKey.Initialize "GetValue debe devolver """" para una clave inexistente"

    Dim config As CConfig
    On Error GoTo TestFail

    ' Arrange
    Set config = New CConfig ' Sin cargar ningún diccionario

    ' Act
    Dim result As String
    result = config.GetValue("NON_EXISTING_KEY")

    ' Assert
    modAssert.AssertEquals "", result, "GetValue debería haber devuelto una cadena vacía."

TestGetValue_ReturnsEmptyForNonExistingKey.Pass
    GoTo Cleanup

TestFail:
    TestGetValue_ReturnsEmptyForNonExistingKey.Fail "Error inesperado: " & Err.Description
Cleanup:
    Set config = Nothing
End Function

' ============================================================================
' PRUEBAS UNITARIAS
' ============================================================================
