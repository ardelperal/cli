VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestSuiteResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Name As String
Public Results As Object ' Clave: Nombre del test, Valor: CTestResult
Public TotalTests As Long
Public PassedTests As Long
Public FailedTests As Long

Private Sub Class_Initialize()
    Set Results = New Scripting.Dictionary
    TotalTests = 0
    PassedTests = 0
    FailedTests = 0
End Sub

' Propiedad que encapsula la l?gica de si la suite ha pasado por completo
Public Property Get AllTestsPassed() As Boolean
    If TotalTests > 0 And FailedTests = 0 Then
        AllTestsPassed = True
    Else
        AllTestsPassed = False
    End If
End Property

' M?todo de inicializaci?n robusto (Lecci?n 21)
Public Sub Initialize(ByVal suiteName As String)
    Name = suiteName
End Sub

' M?todo cr?tico que faltaba
Public Sub AddResult(ByVal testResult As CTestResult)
    If Results.Exists(testResult.Name) Then
        ' Evitar duplicados, aunque no deber?a ocurrir
        Results.Remove testResult.Name
    End If
    
    Results.Add testResult.Name, testResult
    
    TotalTests = TotalTests + 1
    If testResult.Passed Then
        PassedTests = PassedTests + 1
    Else
        FailedTests = FailedTests + 1
    End If
End Sub
