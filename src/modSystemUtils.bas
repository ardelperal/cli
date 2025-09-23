Attribute VB_Name = "modSystemUtils"
Option Compare Database
Option Explicit

Public Sub RunSystemHealthCheck()
    ' Punto de entrada para ejecutar el diagnóstico del sistema desde la UI.
    Dim report As String
    report = modHealthCheck.GenerateHealthReport()
    
    ' Mostrar el informe al usuario.
    Debug.Print report
End Sub
