Attribute VB_Name = "modAssert"
Option Compare Database
Option Explicit


' ============================================================================
' MÓDULO DE ASERCIONES PARA EL FRAMEWORK DE TESTING
' ============================================================================
' Este módulo proporciona funciones de aserción para las pruebas unitarias
' e integración del proyecto CONDOR.

' Códigos de error personalizados
Private Const ERR_ASSERT_TRUE As Long = vbObjectError + 510
Private Const ERR_ASSERT_FALSE As Long = vbObjectError + 511
Private Const ERR_ASSERT_EQUALS As Long = vbObjectError + 512
Private Const ERR_ASSERT_NOT_NULL As Long = vbObjectError + 513
Private Const ERR_ASSERT_IS_NULL As Long = vbObjectError + 514
Private Const ERR_ASSERT_FAIL As Long = vbObjectError + 515
Private Const ERR_ASSERT_NOT_EQUALS As Long = vbObjectError + 516

' ============================================================================
' FUNCIONES DE ASERCIÓN
' ============================================================================

' Verifica que una condición sea verdadera
Public Sub AssertTrue(condition As Boolean, Optional message As String = "")
    If Not condition Then
        Dim errorMsg As String
        errorMsg = "AssertTrue failed"
        If message <> "" Then
            errorMsg = errorMsg & ": " & message
        End If
        Err.Raise ERR_ASSERT_TRUE, "ModAssert.AssertTrue", errorMsg
    End If
End Sub

' Verifica que una condición sea falsa
Public Sub AssertFalse(condition As Boolean, Optional message As String = "")
    If condition Then
        Dim errorMsg As String
        errorMsg = "AssertFalse failed"
        If message <> "" Then
            errorMsg = errorMsg & ": " & message
        End If
        Err.Raise ERR_ASSERT_FALSE, "ModAssert.AssertFalse", errorMsg
    End If
End Sub

' Verifica que dos valores sean iguales
Public Sub AssertEquals(expected As Variant, actual As Variant, Optional message As String = "")
    If expected <> actual Then
        Dim errorMsg As String
        errorMsg = "AssertEquals failed. Expected: '" & CStr(expected) & "', Actual: '" & CStr(actual) & "'"
        If message <> "" Then
            errorMsg = errorMsg & ". " & message
        End If
        Err.Raise ERR_ASSERT_EQUALS, "ModAssert.AssertEquals", errorMsg
    End If
End Sub

' Verifica que dos valores NO sean iguales
Public Sub AssertNotEquals(ByVal value1 As Variant, ByVal value2 As Variant, Optional message As String = "")
    If value1 = value2 Then
        Dim errorMsg As String
        errorMsg = "AssertNotEquals failed. Ambos valores eran: '" & CStr(value1) & "'"
        If message <> "" Then
            errorMsg = errorMsg & ". " & message
        End If
        Err.Raise ERR_ASSERT_NOT_EQUALS, "ModAssert.AssertNotEquals", errorMsg
    End If
End Sub

' Verifica que un objeto no sea Nothing
Public Sub AssertNotNull(obj As Variant, Optional message As String = "")
    If IsNull(obj) Or (IsObject(obj) And obj Is Nothing) Then
        Dim errorMsg As String
        errorMsg = "AssertNotNull failed - object is Nothing or Null"
        If message <> "" Then
            errorMsg = errorMsg & ": " & message
        End If
        Err.Raise ERR_ASSERT_NOT_NULL, "ModAssert.AssertNotNull", errorMsg
    End If
End Sub

' Verifica que un objeto sea Nothing o Null
Public Sub AssertIsNull(obj As Variant, Optional message As String = "")
    Dim isNullOrNothing As Boolean
    isNullOrNothing = IsNull(obj) Or (IsObject(obj) And obj Is Nothing)
    
    If Not isNullOrNothing Then
        Dim errorMsg As String
        errorMsg = "AssertIsNull failed - object is not Nothing or Null"
        If message <> "" Then
            errorMsg = errorMsg & ": " & message
        End If
        Err.Raise ERR_ASSERT_IS_NULL, "ModAssert.AssertIsNull", errorMsg
    End If
End Sub

' Fuerza un fallo en la prueba
Public Sub Fail(Optional message As String = "")
    Dim errorMsg As String
    errorMsg = "Test failed"
    If message <> "" Then
        errorMsg = errorMsg & ": " & message
    End If
    Err.Raise ERR_ASSERT_FAIL, "ModAssert.Fail", errorMsg
End Sub

' ============================================================================
' FUNCIONES DE COMPATIBILIDAD (para mantener compatibilidad con código existente)
' ============================================================================

' Alias para AssertTrue (usado en algunos archivos de prueba)
Public Sub IsTrue(condition As Boolean, Optional message As String = "")
    Call AssertTrue(condition, message)
End Sub

' Alias para AssertEquals (usado en algunos archivos de prueba)
Public Sub AreEqual(expected As Variant, actual As Variant, Optional message As String = "")
    Call AssertEquals(expected, actual, message)
End Sub
