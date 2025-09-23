VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CErrorHandlerService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IErrorHandlerService

Private m_Config As IConfig
Private m_fileSystem As IFileSystem
Private m_logFilePath As String

' ============================================================================
' INICIALIZACI?N CON DEPENDENCIAS (Lecci?n 11)
' ============================================================================
Public Sub Initialize(ByVal config As IConfig, ByVal fileSystem As IFileSystem)
    Set m_Config = config
    Set m_fileSystem = fileSystem
    m_logFilePath = m_Config.GetValue("LOG_FILE_PATH")
End Sub

' ============================================================================
' IMPLEMENTACI?N DE LA INTERFAZ IErrorHandlerService (Lecci?n 1)
' ============================================================================
Private Sub IErrorHandlerService_Initialize(ByVal config As IConfig, ByVal fileSystem As IFileSystem)
    Call Initialize(config, fileSystem)
End Sub

Private Sub IErrorHandlerService_LogError(ByVal errNumber As Long, ByVal errDescription As String, Optional ByVal moduleName As String, Optional ByVal isCritical As Boolean = False)
    On Error Resume Next
    
    Dim logPrefix As String
    If isCritical Then logPrefix = "CRITICAL ERROR" Else logPrefix = "ERROR"
    
    Dim logMessage As String
    logMessage = Now() & " | " & logPrefix & " | " & "Err: " & errNumber & " | " & errDescription & " | Source: " & moduleName
    
    If Len(m_logFilePath) > 0 Then
        m_fileSystem.WriteLineToFile m_logFilePath, logMessage
    Else
        Debug.Print "LogFilePath no configurado. Error no guardado: " & logMessage
    End If
End Sub

Private Sub IErrorHandlerService_LogInfo(ByVal message As String, Optional ByVal moduleName As String)
    ' Implementaci?n futura si es necesario
End Sub

Private Sub IErrorHandlerService_LogWarning(ByVal message As String, Optional ByVal moduleName As String)
    ' Implementaci?n futura si es necesario
End Sub
