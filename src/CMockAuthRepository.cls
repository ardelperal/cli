VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockAuthRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IAuthRepository

' =====================================================
' CLASE: CMockAuthRepository
' PROPOSITO: Mock del repositorio de autenticaci?n para pruebas unitarias
' AUTOR: CONDOR-Expert
' FECHA: 2025-01-15
' =====================================================

' Variables privadas para el mock de GetUserAuthData
Private m_getUserAuthDataReturnValue As EAuthData
Private m_getUserAuthDataWasCalled As Boolean
Private m_lastUserEmailForAuth As String

' ============================================================================
' CONSTRUCTOR
' ============================================================================
Private Sub Class_Initialize()
    Reset
End Sub

' ============================================================================
' M?TODO P?BLICO RESET
' ============================================================================
Public Sub Reset()
    Set m_getUserAuthDataReturnValue = Nothing
    m_getUserAuthDataWasCalled = False
    m_lastUserEmailForAuth = ""
End Sub

' ============================================================================
' M?TODOS DE CONFIGURACI?N PARA TESTS
' ============================================================================

' Configuraci?n del valor de retorno para GetUserAuthData
Public Sub ConfigureGetUserAuthData(ByVal returnValue As EAuthData)
    Set m_getUserAuthDataReturnValue = returnValue
End Sub

' Propiedades "esp?a" para GetUserAuthData
Public Property Get GetUserAuthData_WasCalled() As Boolean
    GetUserAuthData_WasCalled = m_getUserAuthDataWasCalled
End Property

Public Property Get LastUserEmailForAuth() As String
    LastUserEmailForAuth = m_lastUserEmailForAuth
End Property

' ============================================================================
' IMPLEMENTACI?N DE LA INTERFAZ IAuthRepository
' ============================================================================

Private Function IAuthRepository_GetUserAuthData(ByVal userEmail As String) As EAuthData
    m_getUserAuthDataWasCalled = True
    m_lastUserEmailForAuth = userEmail
    
    If Not m_getUserAuthDataReturnValue Is Nothing Then
        Set IAuthRepository_GetUserAuthData = m_getUserAuthDataReturnValue
    Else
        ' Valor por defecto si no se ha configurado
        Dim defaultData As New EAuthData
        defaultData.UserExists = False
        Set IAuthRepository_GetUserAuthData = defaultData
    End If
End Function

' ============================================================================
' M?TODOS P?BLICOS DE CONVENIENCIA - EST?NDAR DE ORO PARA MOCKS
' ============================================================================

Public Function GetUserAuthData(ByVal userEmail As String) As EAuthData
    Set GetUserAuthData = IAuthRepository_GetUserAuthData(userEmail)
End Function
