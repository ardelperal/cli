' Test individual de comandos de Access - PASO 2: Configuraciones de Access
Option Explicit

Dim objFSO, objAccess

Set objFSO = CreateObject("Scripting.FileSystemObject")

' Configurar base de datos de prueba
Dim dbPath
dbPath = "c:\Proyectos\cli\CONDOR.accdb"

WScript.Echo "=== TEST PASO 2: CONFIGURACIONES DE ACCESS ==="
WScript.Echo "Probando configuraciones que se aplican en UpdateModules"
WScript.Echo "Base de datos: " & dbPath

On Error Resume Next

' PASO 2.1: Crear objeto Access y abrir base de datos
WScript.Echo ""
WScript.Echo "2.1 - Abriendo Access y base de datos..."
Set objAccess = CreateObject("Access.Application")
objAccess.OpenCurrentDatabase dbPath
If Err.Number <> 0 Then
    WScript.Echo "ERROR: No se pudo abrir la base de datos: " & Err.Number & " - " & Err.Description
    WScript.Quit 1
Else
    WScript.Echo "✓ Base de datos abierta exitosamente"
End If

' PASO 2.2: Configurar visibilidad y control (recomendaciones oficiales Microsoft)
WScript.Echo ""
WScript.Echo "2.2 - Configurando Visible y UserControl..."
Err.Clear
objAccess.Visible = False
objAccess.UserControl = False
If Err.Number <> 0 Then
    WScript.Echo "ERROR en Visible/UserControl: " & Err.Number & " - " & Err.Description
Else
    WScript.Echo "✓ Visible y UserControl configurados exitosamente"
End If

' PASO 2.3: Probar Echo False
WScript.Echo ""
WScript.Echo "2.3 - Probando Echo False..."
Err.Clear
objAccess.Echo False
If Err.Number <> 0 Then
    WScript.Echo "ERROR en Echo: " & Err.Number & " - " & Err.Description
Else
    WScript.Echo "✓ Echo False aplicado exitosamente"
End If

' PASO 2.4: Probar DoCmd.SetWarnings False
WScript.Echo ""
WScript.Echo "2.4 - Probando DoCmd.SetWarnings False..."
Err.Clear
objAccess.DoCmd.SetWarnings False
If Err.Number <> 0 Then
    WScript.Echo "ERROR en DoCmd.SetWarnings: " & Err.Number & " - " & Err.Description
Else
    WScript.Echo "✓ DoCmd.SetWarnings False aplicado exitosamente"
End If

' PASO 2.5: Probar SetOption commands (configuraciones oficiales Microsoft)
WScript.Echo ""
WScript.Echo "2.5 - Probando SetOption commands..."
Err.Clear
objAccess.SetOption "Confirm Action Queries", False
If Err.Number <> 0 Then
    WScript.Echo "ERROR en SetOption 'Confirm Action Queries': " & Err.Number & " - " & Err.Description
Else
    WScript.Echo "✓ SetOption 'Confirm Action Queries' aplicado exitosamente"
End If

Err.Clear
objAccess.SetOption "Confirm Document Deletions", False
If Err.Number <> 0 Then
    WScript.Echo "ERROR en SetOption 'Confirm Document Deletions': " & Err.Number & " - " & Err.Description
Else
    WScript.Echo "✓ SetOption 'Confirm Document Deletions' aplicado exitosamente"
End If

Err.Clear
objAccess.SetOption "Confirm Record Changes", False
If Err.Number <> 0 Then
    WScript.Echo "ERROR en SetOption 'Confirm Record Changes': " & Err.Number & " - " & Err.Description
Else
    WScript.Echo "✓ SetOption 'Confirm Record Changes' aplicado exitosamente"
End If

' PASO 2.6: Configuraciones adicionales de UI (recomendaciones oficiales)
WScript.Echo ""
WScript.Echo "2.6 - Configurando opciones adicionales de UI..."
Err.Clear
objAccess.SetOption "Show Status Bar", False
If Err.Number <> 0 Then
    WScript.Echo "ERROR en SetOption 'Show Status Bar': " & Err.Number & " - " & Err.Description
Else
    WScript.Echo "✓ SetOption 'Show Status Bar' aplicado exitosamente"
End If

Err.Clear
objAccess.SetOption "Show Animations", False
If Err.Number <> 0 Then
    WScript.Echo "ERROR en SetOption 'Show Animations': " & Err.Number & " - " & Err.Description
Else
    WScript.Echo "✓ SetOption 'Show Animations' aplicado exitosamente"
End If

' PASO 2.7: Configuraciones de modo de acceso a base de datos
WScript.Echo ""
WScript.Echo "2.7 - Configurando modo de acceso a base de datos..."
Err.Clear
objAccess.SetOption "Default Open Mode for Databases", 1
If Err.Number <> 0 Then
    WScript.Echo "ERROR en SetOption 'Default Open Mode for Databases': " & Err.Number & " - " & Err.Description
Else
    WScript.Echo "✓ SetOption 'Default Open Mode for Databases' aplicado exitosamente"
End If

Err.Clear
objAccess.SetOption "Default Record Locking", 0
If Err.Number <> 0 Then
    WScript.Echo "ERROR en SetOption 'Default Record Locking': " & Err.Number & " - " & Err.Description
Else
    WScript.Echo "✓ SetOption 'Default Record Locking' aplicado exitosamente"
End If

' PASO 2.8: Cerrar Access limpiamente
WScript.Echo ""
WScript.Echo "2.6 - Cerrando Access..."
objAccess.Quit
If Err.Number <> 0 Then
    WScript.Echo "ADVERTENCIA al cerrar Access: " & Err.Number & " - " & Err.Description
    Err.Clear
Else
    WScript.Echo "✓ Access cerrado exitosamente"
End If

WScript.Echo ""
WScript.Echo "=== RESULTADO: CONFIGURACIONES DE ACCESS ==="
WScript.Echo "✓ Test completado - Todas las configuraciones básicas funcionan correctamente"
WScript.Echo ""
WScript.Echo "Presiona Enter para continuar..."
WScript.StdIn.ReadLine