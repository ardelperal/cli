VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTestReporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements ITestReporter

Private m_suiteResults As Object

Public Sub Initialize(ByVal suiteResults As Object)
    Set m_suiteResults = suiteResults
End Sub

Private Sub ITestReporter_Initialize(ByVal suiteResults As Object)
    Call Me.Initialize(suiteResults)
End Sub

Private Function ITestReporter_GenerateReport() As String
    ITestReporter_GenerateReport = Me.GenerateReport()
End Function

Public Function GenerateReport() As String
    Dim report As String
    Dim suiteResult As CTestSuiteResult
    Dim testResult As CTestResult
    Dim key As Variant
    
    report = String(80, "=") & vbCrLf
    report = report & "INFORME CONSOLIDADO DE PRUEBAS UNITARIAS" & vbCrLf
    report = report & "Fecha: " & Format(Now, "dd/mm/yyyy hh:nn:ss") & vbCrLf
    report = report & String(80, "=") & vbCrLf & vbCrLf
    
    If m_suiteResults Is Nothing Then
        GenerateReport = "Error: El objeto de resultados de la suite no ha sido inicializado."
        Exit Function
    End If

    For Each key In m_suiteResults.Keys()
        Set suiteResult = m_suiteResults(key)
        report = report & "SUITE: " & suiteResult.Name & " (" & suiteResult.PassedTests & "/" & suiteResult.TotalTests & " pasadas)" & vbCrLf
        report = report & String(Len("SUITE: " & suiteResult.Name), "-") & vbCrLf
        
        Dim testKey As Variant
        For Each testKey In suiteResult.Results.Keys()
            Set testResult = suiteResult.Results(testKey)
            If testResult.Passed Then
                report = report & " [V] PASS: " & testResult.Name & vbCrLf
            Else
                report = report & " [X] FAIL: " & testResult.Name & vbCrLf
                report = report & "     L--> Error: " & testResult.ErrorMessage & vbCrLf
            End If
        Next testKey
        report = report & vbCrLf
    Next key
    
    report = report & String(80, "-") & vbCrLf
    report = report & "RESUMEN GLOBAL" & vbCrLf
    report = report & "Suites ejecutadas: " & Me.TotalSuites & vbCrLf
    report = report & "Total de pruebas: " & Me.TotalTests & vbCrLf
    report = report & "Pruebas exitosas: " & Me.TotalPassed & vbCrLf
    report = report & "Pruebas fallidas: " & Me.TotalFailed & vbCrLf
    report = report & String(80, "-") & vbCrLf
    
    If Me.TotalFailed = 0 Then
        report = report & "RESULTADO: V TODAS LAS PRUEBAS PASARON" & vbCrLf
    Else
        report = report & "RESULTADO: X HAY PRUEBAS FALLIDAS" & vbCrLf
    End If
    
    report = report & String(80, "=") & vbCrLf
    
    GenerateReport = report
End Function

Public Property Get TotalSuites() As Long
    If m_suiteResults Is Nothing Then
        TotalSuites = 0
    Else
        TotalSuites = m_suiteResults.Count
    End If
End Property

Public Property Get TotalTests() As Long
    Dim total As Long
    Dim key As Variant
    Dim suiteResult As CTestSuiteResult
    
    total = 0
    If Not m_suiteResults Is Nothing Then
        For Each key In m_suiteResults.Keys()
            Set suiteResult = m_suiteResults(key)
            total = total + suiteResult.TotalTests
        Next key
    End If
    TotalTests = total
End Property

Public Property Get TotalPassed() As Long
    Dim total As Long
    Dim key As Variant
    Dim suiteResult As CTestSuiteResult
    
    total = 0
    If Not m_suiteResults Is Nothing Then
        For Each key In m_suiteResults.Keys()
            Set suiteResult = m_suiteResults(key)
            total = total + suiteResult.PassedTests
        Next key
    End If
    TotalPassed = total
End Property

Public Property Get TotalFailed() As Long
    Dim total As Long
    Dim key As Variant
    Dim suiteResult As CTestSuiteResult
    
    total = 0
    If Not m_suiteResults Is Nothing Then
        For Each key In m_suiteResults.Keys()
            Set suiteResult = m_suiteResults(key)
            total = total + suiteResult.FailedTests
        Next key
    End If
    TotalFailed = total
End Property
