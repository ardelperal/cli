VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockOperationRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IOperationRepository

' Variables esp?a privadas (Lecci?n 22)
Private m_SaveLog_WasCalled As Boolean
Private m_CallCount As Long
Private m_SaveLog_LastEntry As EOperationLog

' Propiedades p?blicas de solo lectura para verificaci?n en tests
Public Property Get SaveLog_WasCalled() As Boolean
    SaveLog_WasCalled = m_SaveLog_WasCalled
End Property
Public Property Get CallCount() As Long
    CallCount = m_CallCount
End Property
Public Property Get SaveLog_LastEntry() As EOperationLog
    Set SaveLog_LastEntry = m_SaveLog_LastEntry
End Property

' Constructor
Private Sub Class_Initialize()
    Call Reset
End Sub

Private Sub IOperationRepository_SaveLog(ByVal logEntry As EOperationLog)
    ' Registrar la llamada para verificaci?n en tests
    m_SaveLog_WasCalled = True
    m_CallCount = m_CallCount + 1
    Set m_SaveLog_LastEntry = logEntry
End Sub

' Resetea el estado del mock para asegurar el aislamiento entre tests
Public Sub Reset()
    m_SaveLog_WasCalled = False
    m_CallCount = 0
    Set m_SaveLog_LastEntry = Nothing
End Sub
