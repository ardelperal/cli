Option Explicit

' Script para probar configuraciones sin DisplayAlerts
' que causa Error 438: El objeto no acepta esta propiedad o método

WScript.Echo "=== Probando configuraciones SIN DisplayAlerts ==="

Dim app
Set app = CreateObject("Access.Application")

On Error Resume Next
Err.Clear

WScript.Echo "1. Configurando Visible = False..."
app.Visible = False
If Err.Number <> 0 Then
    WScript.Echo "ERROR: " & Err.Number & ": " & Err.Description
    Err.Clear
Else
    WScript.Echo "OK"
End If

WScript.Echo "2. Configurando UserControl = False..."
app.UserControl = False
If Err.Number <> 0 Then
    WScript.Echo "ERROR: " & Err.Number & ": " & Err.Description
    Err.Clear
Else
    WScript.Echo "OK"
End If

WScript.Echo "3. Abriendo base de datos..."
app.OpenCurrentDatabase "c:\Proyectos\cli\CONDOR.accdb"
If Err.Number <> 0 Then
    WScript.Echo "ERROR abriendo DB: " & Err.Number & ": " & Err.Description
    app.Quit 1
    Set app = Nothing
    WScript.Quit 1
Else
    WScript.Echo "OK"
End If

WScript.Echo "4. Configurando Echo False..."
app.Echo False
If Err.Number <> 0 Then
    WScript.Echo "ERROR: " & Err.Number & ": " & Err.Description
    Err.Clear
Else
    WScript.Echo "OK"
End If

WScript.Echo "5. Configurando SetWarnings False..."
app.DoCmd.SetWarnings False
If Err.Number <> 0 Then
    WScript.Echo "ERROR: " & Err.Number & ": " & Err.Description
    Err.Clear
Else
    WScript.Echo "OK"
End If

WScript.Echo "6. Configurando VBE.MainWindow.Visible = False..."
app.VBE.MainWindow.Visible = False
If Err.Number <> 0 Then
    WScript.Echo "ERROR: " & Err.Number & ": " & Err.Description
    Err.Clear
Else
    WScript.Echo "OK"
End If

WScript.Echo "7. Configurando opciones de aplicación..."
app.Application.SetOption "Confirm Action Queries", False
If Err.Number <> 0 Then
    WScript.Echo "ERROR en Confirm Action Queries: " & Err.Number & ": " & Err.Description
    Err.Clear
Else
    WScript.Echo "OK - Confirm Action Queries"
End If

app.Application.SetOption "Confirm Document Deletions", False
If Err.Number <> 0 Then
    WScript.Echo "ERROR en Confirm Document Deletions: " & Err.Number & ": " & Err.Description
    Err.Clear
Else
    WScript.Echo "OK - Confirm Document Deletions"
End If

app.Application.SetOption "Confirm Record Changes", False
If Err.Number <> 0 Then
    WScript.Echo "ERROR en Confirm Record Changes: " & Err.Number & ": " & Err.Description
    Err.Clear
Else
    WScript.Echo "OK - Confirm Record Changes"
End If

On Error GoTo 0

WScript.Echo ""
WScript.Echo "=== TODAS LAS CONFIGURACIONES APLICADAS SIN DisplayAlerts ==="
WScript.Echo "Esperando 3 segundos para verificar que no aparece UI..."
WScript.Sleep 3000

WScript.Echo "Cerrando Access..."
app.Quit 1
Set app = Nothing

WScript.Echo "✓ Prueba completada exitosamente"