Option Explicit

' Script de test para verificar el contenido de módulos antes y después del rebuild
' Uso: cscript test_module_content.vbs [nombre_modulo]

Dim objFSO, objAccess, moduleName, testResults
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Obtener el nombre del módulo a testear (por defecto CAppManager)
If WScript.Arguments.Count > 0 Then
    moduleName = WScript.Arguments(0)
Else
    moduleName = "CAppManager"
End If

WScript.Echo "=== TEST DE CONTENIDO DE MÓDULO: " & moduleName & " ==="
WScript.Echo ""

' Paso 1: Verificar contenido del archivo fuente
WScript.Echo "1. Verificando archivo fuente..."
Dim sourceFile, sourceContent
sourceFile = "C:\Proyectos\cli\src\" & moduleName & ".cls"

If Not objFSO.FileExists(sourceFile) Then
    ' Intentar con extensión .bas si no se encuentra .cls
    sourceFile = "C:\Proyectos\cli\src\" & moduleName & ".bas"
    If Not objFSO.FileExists(sourceFile) Then
        WScript.Echo "❌ ERROR: No se encuentra el archivo fuente: " & sourceFile
        WScript.Quit 1
    End If
End If

Dim file, textStream
Set textStream = objFSO.OpenTextFile(sourceFile, 1, False, -1) ' -1 = Unicode
sourceContent = textStream.ReadAll
textStream.Close

WScript.Echo "✓ Archivo fuente encontrado: " & sourceFile
WScript.Echo "✓ Tamaño del contenido: " & Len(sourceContent) & " caracteres"
WScript.Echo "✓ Líneas en el archivo: " & (Len(sourceContent) - Len(Replace(sourceContent, vbCrLf, ""))) / 2 + 1

' Mostrar las primeras líneas del archivo fuente
Dim sourceLines
sourceLines = Split(sourceContent, vbCrLf)
WScript.Echo "✓ Primeras 5 líneas del archivo fuente:"
Dim i
For i = 0 To 4
    If i < UBound(sourceLines) + 1 Then
        WScript.Echo "   " & (i + 1) & ": " & sourceLines(i)
    End If
Next

WScript.Echo ""

' Paso 2: Abrir Access y verificar contenido del módulo en VBE
WScript.Echo "2. Verificando módulo en VBE..."

On Error Resume Next
Set objAccess = CreateObject("Access.Application")
If Err.Number <> 0 Then
    WScript.Echo "❌ ERROR: No se pudo crear la aplicación Access: " & Err.Description
    WScript.Quit 1
End If
Err.Clear

' Abrir la base de datos
objAccess.OpenCurrentDatabase "C:\Proyectos\cli\CONDOR.accdb"
If Err.Number <> 0 Then
    WScript.Echo "❌ ERROR: No se pudo abrir la base de datos: " & Err.Description
    objAccess.Quit
    WScript.Quit 1
End If

' Acceder al VBE
Dim vbProject, vbComponent, moduleFound
Set vbProject = objAccess.VBE.ActiveVBProject
moduleFound = False

For Each vbComponent In vbProject.VBComponents
    If vbComponent.Name = moduleName Then
        moduleFound = True
        WScript.Echo "✓ Módulo encontrado en VBE: " & moduleName
        
        ' Verificar contenido del módulo
        Dim codeModule, moduleContent, moduleLines
        Set codeModule = vbComponent.CodeModule
        
        If codeModule.CountOfLines > 0 Then
            moduleContent = codeModule.Lines(1, codeModule.CountOfLines)
            moduleLines = codeModule.CountOfLines
            
            WScript.Echo "✓ Líneas en el módulo VBE: " & moduleLines
            WScript.Echo "✓ Tamaño del contenido VBE: " & Len(moduleContent) & " caracteres"
            
            ' Mostrar las primeras líneas del módulo en VBE
            WScript.Echo "✓ Primeras 5 líneas del módulo en VBE:"
            Dim vbeLines
            vbeLines = Split(moduleContent, vbCrLf)
            For i = 0 To 4
                If i < UBound(vbeLines) + 1 Then
                    WScript.Echo "   " & (i + 1) & ": " & vbeLines(i)
                End If
            Next
            
            ' Comparar contenidos
            WScript.Echo ""
            WScript.Echo "3. Comparando contenidos..."
            
            ' Normalizar contenidos para comparación (eliminar espacios extra, etc.)
            Dim normalizedSource, normalizedVBE
            normalizedSource = Trim(Replace(Replace(sourceContent, vbCrLf, vbLf), vbCr, vbLf))
            normalizedVBE = Trim(Replace(Replace(moduleContent, vbCrLf, vbLf), vbCr, vbLf))
            
            If normalizedSource = normalizedVBE Then
                WScript.Echo "✅ ÉXITO: El contenido del archivo fuente coincide con el módulo en VBE"
            Else
                WScript.Echo "❌ PROBLEMA: El contenido NO coincide"
                WScript.Echo "   - Archivo fuente: " & Len(normalizedSource) & " caracteres"
                WScript.Echo "   - Módulo VBE: " & Len(normalizedVBE) & " caracteres"
                
                ' Mostrar diferencias básicas
                If Len(normalizedVBE) < 50 Then
                    WScript.Echo "   - Contenido VBE parece estar vacío o casi vacío"
                    WScript.Echo "   - Contenido completo VBE: [" & moduleContent & "]"
                End If
            End If
        Else
            WScript.Echo "❌ PROBLEMA: El módulo en VBE está completamente vacío (0 líneas)"
        End If
        
        Exit For
    End If
Next

If Not moduleFound Then
    WScript.Echo "❌ ERROR: Módulo " & moduleName & " no encontrado en VBE"
End If

' Cerrar Access
objAccess.CloseCurrentDatabase
objAccess.Quit
Set objAccess = Nothing

WScript.Echo ""
WScript.Echo "=== FIN DEL TEST ==="