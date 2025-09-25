Option Explicit

' Script exhaustivo para probar comandos de Access CON errores visibles
' OBJETIVO: Detectar errores silenciosos permitiendo que Access "se queje"

WScript.Echo "=== DIAGNÓSTICO EXHAUSTIVO DE COMANDOS ACCESS ==="
WScript.Echo "IMPORTANTE: Este script PERMITIRÁ que aparezca la UI de Access"
WScript.Echo "para detectar errores que normalmente están silenciados."
WScript.Echo ""
WScript.Echo "Presiona ENTER para continuar o Ctrl+C para cancelar..."
WScript.StdIn.ReadLine

Dim app, testResults, testCount
Set testResults = CreateObject("Scripting.Dictionary")
testCount = 0

' Función para registrar resultado de prueba
Sub LogTest(testName, success, errorMsg)
    testCount = testCount + 1
    WScript.Echo testCount & ". " & testName & ": " & IIf(success, "✓ OK", "✗ ERROR - " & errorMsg)
    testResults.Add testName, Array(success, errorMsg)
End Sub

' Función para esperar y observar UI
Sub WaitAndObserve(seconds, description)
    WScript.Echo "   → Esperando " & seconds & "s para observar: " & description
    WScript.Sleep seconds * 1000
End Sub

WScript.Echo "=== FASE 1: CREACIÓN Y CONFIGURACIÓN BÁSICA ==="

' Test 1: Crear instancia de Access (CON UI visible)
WScript.Echo "1. Creando instancia de Access (UI VISIBLE)..."
On Error Resume Next
Err.Clear
Set app = CreateObject("Access.Application")
If Err.Number <> 0 Then
    LogTest "Crear Access.Application", False, Err.Description
    WScript.Quit 1
Else
    LogTest "Crear Access.Application", True, ""
End If

' Test 2: Configurar Visible = True (para ver errores)
app.Visible = True
app.UserControl = True
LogTest "Configurar Visible=True, UserControl=True", Err.Number = 0, Err.Description
WaitAndObserve 2, "Access debe estar visible"

' Test 3: Abrir base de datos
WScript.Echo ""
WScript.Echo "=== FASE 2: APERTURA DE BASE DE DATOS ==="
app.OpenCurrentDatabase "c:\Proyectos\cli\CONDOR.accdb"
LogTest "Abrir base de datos", Err.Number = 0, Err.Description
WaitAndObserve 3, "Base de datos abierta"

' Test 4: Probar diferentes sintaxis de Echo
WScript.Echo ""
WScript.Echo "=== FASE 3: COMANDOS ECHO ==="
Err.Clear
app.Echo False
LogTest "app.Echo False", Err.Number = 0, Err.Description

Err.Clear
app.Application.Echo False
LogTest "app.Application.Echo False", Err.Number = 0, Err.Description

' Test 5: Probar SetWarnings
WScript.Echo ""
WScript.Echo "=== FASE 4: COMANDOS SETWARNINGS ==="
Err.Clear
app.DoCmd.SetWarnings False
LogTest "app.DoCmd.SetWarnings False", Err.Number = 0, Err.Description

' Test 6: Probar SetOption - diferentes opciones
WScript.Echo ""
WScript.Echo "=== FASE 5: COMANDOS SETOPTION ==="
Dim options : options = Array( _
    "Confirm Action Queries", _
    "Confirm Document Deletions", _
    "Confirm Record Changes", _
    "Confirm Report Close", _
    "Status Bar Text", _
    "Show Status Bar" _
)

Dim i
For i = 0 To UBound(options)
    Err.Clear
    app.Application.SetOption options(i), False
    LogTest "SetOption '" & options(i) & "'", Err.Number = 0, Err.Description
    If Err.Number <> 0 Then
        WaitAndObserve 1, "Error en SetOption - observar diálogos"
    End If
Next

' Test 7: Probar AutomationSecurity
WScript.Echo ""
WScript.Echo "=== FASE 6: AUTOMATION SECURITY ==="
Err.Clear
app.Application.AutomationSecurity = 1
LogTest "AutomationSecurity = 1", Err.Number = 0, Err.Description

' Test 8: Probar acceso a VBE
WScript.Echo ""
WScript.Echo "=== FASE 7: ACCESO A VBE ==="
Err.Clear
Dim vbeAccess : vbeAccess = Not (app.VBE Is Nothing)
LogTest "Acceso a app.VBE", vbeAccess And Err.Number = 0, Err.Description

If vbeAccess Then
    Err.Clear
    app.VBE.MainWindow.Visible = True  ' Hacer visible para ver errores
    LogTest "VBE.MainWindow.Visible = True", Err.Number = 0, Err.Description
    WaitAndObserve 2, "Ventana VBE debe estar visible"
    
    Err.Clear
    app.VBE.MainWindow.Visible = False
    LogTest "VBE.MainWindow.Visible = False", Err.Number = 0, Err.Description
End If

' Test 9: Probar operaciones con módulos VBA
WScript.Echo ""
WScript.Echo "=== FASE 8: OPERACIONES VBA ==="
If vbeAccess Then
    Err.Clear
    Dim vbProject : Set vbProject = app.VBE.ActiveVBProject
    LogTest "Acceso a ActiveVBProject", Not (vbProject Is Nothing) And Err.Number = 0, Err.Description
    
    If Not (vbProject Is Nothing) Then
        ' Contar módulos existentes
        Err.Clear
        Dim moduleCount : moduleCount = vbProject.VBComponents.Count
        LogTest "Contar módulos VBA (" & moduleCount & ")", Err.Number = 0, Err.Description
        
        ' Intentar crear un módulo de prueba
        Err.Clear
        Dim testModule
        Set testModule = vbProject.VBComponents.Add(1) ' vbext_ct_StdModule
        LogTest "Crear módulo de prueba", Err.Number = 0, Err.Description
        WaitAndObserve 2, "Módulo creado - observar cambios"
        
        If Err.Number = 0 And Not (testModule Is Nothing) Then
            ' Agregar código al módulo
            Err.Clear
            testModule.CodeModule.AddFromString "Sub TestSub()" & vbCrLf & "End Sub"
            LogTest "Agregar código al módulo", Err.Number = 0, Err.Description
            
            ' Eliminar el módulo de prueba
            Err.Clear
            vbProject.VBComponents.Remove testModule
            LogTest "Eliminar módulo de prueba", Err.Number = 0, Err.Description
            WaitAndObserve 2, "Módulo eliminado - observar cambios"
        End If
    End If
End If

' Test 10: Probar compilación
WScript.Echo ""
WScript.Echo "=== FASE 9: COMPILACIÓN ==="
Err.Clear
app.DoCmd.RunCommand 126  ' acCmdCompileAndSaveAllModules
LogTest "Compilar y guardar módulos", Err.Number = 0, Err.Description
WaitAndObserve 3, "Compilación - observar mensajes"

' Test 11: Probar importación de módulo real
WScript.Echo ""
WScript.Echo "=== FASE 10: IMPORTACIÓN REAL ==="
If vbeAccess Then
    Err.Clear
    app.DoCmd.TransferText 0, , "TestImport", "c:\Proyectos\cli\src\CMockConfig.cls"
    LogTest "Importar módulo real (TransferText)", Err.Number = 0, Err.Description
    
    ' Método alternativo de importación
    Err.Clear
    vbProject.VBComponents.Import "c:\Proyectos\cli\src\CMockConfig.cls"
    LogTest "Importar módulo real (VBComponents.Import)", Err.Number = 0, Err.Description
    WaitAndObserve 3, "Importación - observar diálogos"
End If

' Test 12: Operaciones de guardado
WScript.Echo ""
WScript.Echo "=== FASE 11: OPERACIONES DE GUARDADO ==="
Err.Clear
app.DoCmd.Save
LogTest "DoCmd.Save", Err.Number = 0, Err.Description

Err.Clear
app.CurrentProject.Connection.Execute "SELECT 1"  ' Prueba de conexión
LogTest "Prueba de conexión BD", Err.Number = 0, Err.Description

On Error GoTo 0

WScript.Echo ""
WScript.Echo "=== RESUMEN FINAL ==="
WScript.Echo "Total de pruebas: " & testResults.Count
Dim successCount : successCount = 0
Dim key
For Each key In testResults.Keys
    If testResults(key)(0) Then successCount = successCount + 1
Next

WScript.Echo "Exitosas: " & successCount
WScript.Echo "Fallidas: " & (testResults.Count - successCount)

WScript.Echo ""
WScript.Echo "=== ERRORES DETECTADOS ==="
For Each key In testResults.Keys
    If Not testResults(key)(0) Then
        WScript.Echo "✗ " & key & ": " & testResults(key)(1)
    End If
Next

WScript.Echo ""
WScript.Echo "Presiona ENTER para cerrar Access..."
WScript.StdIn.ReadLine

' Cerrar Access
app.Quit 1
Set app = Nothing

WScript.Echo "✓ Diagnóstico completado"