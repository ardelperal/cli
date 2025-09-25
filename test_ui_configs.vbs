Option Explicit

' Script para probar cada configuración de UI individualmente
' y detectar cuál causa errores que hacen aparecer la interfaz

Dim objAccess, testResults, i
Set testResults = CreateObject("Scripting.Dictionary")

' Función para crear una instancia limpia de Access
Function CreateCleanAccess()
    Set CreateCleanAccess = CreateObject("Access.Application")
End Function

' Función para cerrar Access de forma segura
Sub CloseAccess(app)
    On Error Resume Next
    If Not app Is Nothing Then
        app.Quit 1
        Set app = Nothing
    End If
    On Error GoTo 0
End Sub

' Función para probar una configuración específica
Function TestConfiguration(configName, configCode)
    Dim app, errorOccurred, errorDesc
    errorOccurred = False
    errorDesc = ""
    
    WScript.Echo "=== Probando: " & configName & " ==="
    
    Set app = CreateCleanAccess()
    
    On Error Resume Next
    Err.Clear
    
    ' Ejecutar el código de configuración
    Execute configCode
    
    If Err.Number <> 0 Then
        errorOccurred = True
        errorDesc = "Error " & Err.Number & ": " & Err.Description
        WScript.Echo "ERROR: " & errorDesc
    Else
        WScript.Echo "OK: Configuración aplicada sin errores"
    End If
    
    On Error GoTo 0
    
    ' Esperar un momento para ver si aparece UI
    WScript.Sleep 1000
    
    CloseAccess app
    
    TestConfiguration = Not errorOccurred
    testResults.Add configName, Array(Not errorOccurred, errorDesc)
End Function

' Función para probar configuración con base de datos abierta
Function TestConfigurationWithDB(configName, configCode)
    Dim app, errorOccurred, errorDesc
    errorOccurred = False
    errorDesc = ""
    
    WScript.Echo "=== Probando con DB: " & configName & " ==="
    
    Set app = CreateCleanAccess()
    
    On Error Resume Next
    Err.Clear
    
    ' Primero abrir la base de datos
    app.OpenCurrentDatabase "c:\Proyectos\cli\CONDOR.accdb"
    
    If Err.Number <> 0 Then
        WScript.Echo "ERROR abriendo DB: " & Err.Number & ": " & Err.Description
        CloseAccess app
        TestConfigurationWithDB = False
        Exit Function
    End If
    
    ' Luego ejecutar el código de configuración
    Execute configCode
    
    If Err.Number <> 0 Then
        errorOccurred = True
        errorDesc = "Error " & Err.Number & ": " & Err.Description
        WScript.Echo "ERROR: " & errorDesc
    Else
        WScript.Echo "OK: Configuración aplicada sin errores"
    End If
    
    On Error GoTo 0
    
    ' Esperar un momento para ver si aparece UI
    WScript.Sleep 1000
    
    CloseAccess app
    
    TestConfigurationWithDB = Not errorOccurred
    testResults.Add configName & "_WithDB", Array(Not errorOccurred, errorDesc)
End Function

' INICIO DE PRUEBAS
WScript.Echo "Iniciando pruebas de configuraciones de UI..."
WScript.Echo "Presiona Ctrl+C si ves que aparece la interfaz de Access"
WScript.Echo ""

' Prueba 1: objApp.Visible = False
TestConfiguration "Visible_False", "app.Visible = False"

' Prueba 2: objApp.UserControl = False  
TestConfiguration "UserControl_False", "app.UserControl = False"

' Prueba 3: objApp.Echo False
TestConfigurationWithDB "Echo_False", "app.Echo False"

' Prueba 4: objApp.DoCmd.SetWarnings False
TestConfigurationWithDB "SetWarnings_False", "app.DoCmd.SetWarnings False"

' Prueba 5: objApp.DisplayAlerts = False
TestConfiguration "DisplayAlerts_False", "app.DisplayAlerts = False"

' Prueba 6: objApp.VBE.MainWindow.Visible = False
TestConfigurationWithDB "VBE_MainWindow_False", "app.VBE.MainWindow.Visible = False"

' Prueba 7: Combinación de configuraciones básicas
TestConfiguration "Basic_Combo", "app.Visible = False: app.UserControl = False: app.DisplayAlerts = False"

' Prueba 8: Todas las configuraciones juntas con DB
TestConfigurationWithDB "All_Configs", "app.Visible = False: app.UserControl = False: app.Echo False: app.DoCmd.SetWarnings False: app.DisplayAlerts = False: app.VBE.MainWindow.Visible = False"

' RESUMEN DE RESULTADOS
WScript.Echo ""
WScript.Echo "=== RESUMEN DE RESULTADOS ==="
Dim key
For Each key In testResults.Keys
    Dim result : result = testResults(key)
    If result(0) Then
        WScript.Echo "✓ " & key & ": OK"
    Else
        WScript.Echo "✗ " & key & ": FALLO - " & result(1)
    End If
Next

WScript.Echo ""
WScript.Echo "Pruebas completadas. Revisa si alguna configuración hizo aparecer la UI."