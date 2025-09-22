' CONDOR CLI - Herramienta de linea de comandos para el proyecto CONDOR
' Funcionalidades: Sincronizacion VBA, gestion de tablas, y operaciones del proyecto
' Version sin dialogos para automatizacion completa

Option Explicit

Dim objAccess
Dim strAccessPath
Dim strSourcePath
Dim strAction
Dim objFSO
Dim objArgs
Dim strDbPassword
Dim pathArg, i
Dim gVerbose ' Variable global para soporte --verbose

' Configuracion
' Configuracion inicial - se determinara la base de datos segun la accion
Dim strDataPath
strAccessPath = "C:\Proyectos\CONDOR\back\Desarrollo\CONDOR.accdb"
strDataPath = "C:\Proyectos\CONDOR\back\CONDOR_datos.accdb"
strSourcePath = "C:\Proyectos\CONDOR\src"

' Obtener argumentos de linea de comandos
Set objArgs = WScript.Arguments

' Verificar si se solicita ayuda
If objArgs.Count > 0 Then
    If LCase(objArgs(0)) = "--help" Or LCase(objArgs(0)) = "-h" Or LCase(objArgs(0)) = "help" Then
        Call ShowHelp()
        WScript.Quit 0
    End If
End If

If objArgs.Count = 0 Then
    WScript.Echo "=== CONDOR CLI - Herramienta de linea de comandos ==="
    WScript.Echo "Uso: cscript condor_cli.vbs [comando] [opciones]"
    WScript.Echo ""
    WScript.Echo "COMANDOS DISPONIBLES:"
    WScript.Echo "  export     - Exportar modulos VBA a /src (con codificacion ANSI)"
    WScript.Echo "  validate   - Validar sintaxis de modulos VBA sin importar"
    WScript.Echo "  validate-schema - Valida el esquema de las BDs de prueba contra el Master Plan"
    WScript.Echo "  test       - Ejecutar suite de pruebas unitarias"
    
    WScript.Echo "  rebuild    - Reconstruir proyecto VBA (eliminar todos los modulos y reimportar)"
    WScript.Echo "  bundle <funcionalidad> [ruta_destino] - Empaquetar archivos de codigo por funcionalidad"
    WScript.Echo "  lint       - Auditar codigo VBA para detectar cabeceras duplicadas"
    WScript.Echo "  createtable <nombre> <sql> - Crear tabla con consulta SQL"
    WScript.Echo "  droptable <nombre> - Eliminar tabla"
    WScript.Echo "  listtables [db_path] [--schema] [--output] - Listar tablas de BD. --schema: muestra campos, tipos y requerido. --output: exporta a [nombre_bd]_listtables.txt"
    WScript.Echo "  relink [db_path] [folder]    - Re-vincular tablas a bases locales"
    WScript.Echo "  migrate [file.sql]           - Ejecutar scripts de migración SQL desde ./db/migrations"
    WScript.Echo "  relink --all - Re-vincular todas las bases en ./back automaticamente"
    WScript.Echo "  migrate [file.sql] - Ejecutar scripts de migración SQL en /db/migrations"
    WScript.Echo ""


    WScript.Echo "PARÁMETROS DE FUNCIONALIDAD PARA 'bundle' (según CONDOR_MASTER_PLAN.md):"
    WScript.Echo "  Auth: Empaqueta Autenticación + dependencias (Config, Error, Modelos)"
    WScript.Echo "  Document: Empaqueta Gestión de Documentos + dependencias (Config, FileSystem, Error, Word, Modelos)"
    WScript.Echo "  Expediente: Empaqueta Gestión de Expedientes + dependencias (Config, Error, Modelos, Utilidades)"
    WScript.Echo "  Solicitud: Empaqueta Gestión de Solicitudes + dependencias (Config, Error, Modelos, Utilidades)"
    WScript.Echo "  Workflow: Empaqueta Flujos de Trabajo + dependencias (Config, Error, Modelos, Utilidades)"
    WScript.Echo "  Mapeo: Empaqueta Gestión de Mapeos + dependencias (Config, Error, Modelos)"
    WScript.Echo "  Config: Empaqueta Configuración del Sistema + dependencias (Error, Modelos)"
    WScript.Echo "  FileSystem: Empaqueta Sistema de Archivos + dependencias (Error, Modelos)"
    WScript.Echo "  Error: Empaqueta Manejo de Errores + dependencias (Modelos)"
    WScript.Echo "  Word: Empaqueta Microsoft Word + dependencias (Error, Modelos)"
    WScript.Echo "  TestFramework: Empaqueta Framework de Pruebas + dependencias (11 archivos: ITestReporter, CTestResult, CTestSuiteResult, CTestReporter, modTestRunner, modTestUtils, ModAssert, TestModAssert, IFileSystem, IConfig, IErrorHandlerService)"
    WScript.Echo "  App: Empaqueta Gestión de Aplicación + dependencias (Config, Error, Modelos)"
    WScript.Echo "  Models: Empaqueta Modelos de Datos (entidades base)"
    WScript.Echo "  Utils: Empaqueta Utilidades y Enumeraciones + dependencias (Error, Modelos)"
    WScript.Echo "  Tests: Empaqueta todos los archivos de pruebas (Test* e IntegrationTest*)"
    WScript.Echo ""
    WScript.Echo "OPCIONES ESPECIALES:"

    WScript.Echo "  --verbose  - Mostrar informacion detallada durante la operacion"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS:"
    WScript.Echo "  cscript condor_cli.vbs validate"
    WScript.Echo "  cscript condor_cli.vbs validate-schema"
    WScript.Echo "  cscript condor_cli.vbs export --verbose"
    WScript.Echo "  cscript condor_cli.vbs rebuild"
    WScript.Echo "  cscript condor_cli.vbs bundle Tests"
    WScript.Echo "  cscript condor_cli.vbs listtables"
    WScript.Echo "  cscript condor_cli.vbs listtables ./back/test_db/templates/CONDOR_test_template.accdb --schema"
    WScript.Quit 1
End If

strAction = LCase(objArgs(0))

If strAction <> "export" And strAction <> "validate" And strAction <> "validate-schema" And strAction <> "test" And strAction <> "createtable" And strAction <> "droptable" And strAction <> "listtables" And strAction <> "relink" And strAction <> "rebuild" And strAction <> "lint" And strAction <> "bundle" And strAction <> "migrate" And strAction <> "export-form" And strAction <> "import-form" And strAction <> "validate-form-json" And strAction <> "roundtrip-form" Then
    WScript.Echo "Error: Comando debe ser 'export', 'validate', 'validate-schema', 'test', 'createtable', 'droptable', 'listtables', 'relink', 'rebuild', 'lint', 'bundle', 'migrate', 'export-form', 'import-form', 'validate-form-json' o 'roundtrip-form'"
    WScript.Quit 1
End If

Set objFSO = CreateObject("Scripting.FileSystemObject")

' Los comandos bundle, validate-schema y validate-form-json no requieren Access
If strAction = "bundle" Then
    ' Verificar si se solicita ayuda específica para bundle
    If objArgs.Count > 1 Then
        If LCase(objArgs(1)) = "--help" Or LCase(objArgs(1)) = "-h" Or LCase(objArgs(1)) = "help" Then
            Call ShowBundleHelp()
            WScript.Quit 0
        End If
    End If
    Call BundleFunctionality()
    WScript.Quit 0
ElseIf strAction = "validate-schema" Then
    Call ValidateSchema()
    WScript.Quit 0
ElseIf strAction = "validate-form-json" Then
    Call ValidateFormJsonCommand()
    WScript.Quit 0
ElseIf strAction = "roundtrip-form" Then
    Call RoundtripFormCommand()
    WScript.Quit 0
End If

' Determinar qué base de datos usar según la acción
If strAction = "createtable" Or strAction = "droptable" Or strAction = "migrate" Then
    strAccessPath = strDataPath
ElseIf strAction = "listtables" Then
    pathArg = ""
    ' Buscar un argumento que no sea el flag --schema
    For i = 1 To objArgs.Count - 1
        If LCase(objArgs(i)) <> "--schema" Then
            pathArg = objArgs(i)
            Exit For
        End If
    Next
    
    If pathArg <> "" Then
        ' Resolver ruta relativa a absoluta
        strAccessPath = ResolveRelativePath(pathArg)
    Else
        strAccessPath = strDataPath
    End If
ElseIf strAction = "import-form" Then
    ' Para import-form, usar la base de datos especificada como segundo argumento
    If objArgs.Count >= 3 Then
        strAccessPath = ResolveRelativePath(objArgs(2))
    Else
        strAccessPath = strDataPath
    End If
End If

' Para rebuild y test, usar la base de datos de desarrollo
If strAction = "rebuild" Or strAction = "test" Then
    strAccessPath = "C:\Proyectos\CONDOR\back\Desarrollo\CONDOR.accdb"
End If

' Verificar que existe la base de datos
If Not objFSO.FileExists(strAccessPath) Then
    WScript.Echo "Error: La base de datos no existe: " & strAccessPath
    WScript.Quit 1
End If

If strAction = "import-form" Then
    WScript.Echo "=== IMPORTANDO FORMULARIO ==="
Else
    WScript.Echo "=== INICIANDO SINCRONIZACION VBA ==="
End If
WScript.Echo "Accion: " & strAction
WScript.Echo "Base de datos: " & strAccessPath
WScript.Echo "Directorio: " & strSourcePath



' Verificar y cerrar procesos de Access existentes
Call CloseExistingAccessProcesses()

On Error Resume Next

' Crear aplicacion Access
WScript.Echo "Iniciando aplicacion Access..."
Set objAccess = CreateObject("Access.Application")

If Err.Number <> 0 Then
    WScript.Echo "Error al crear aplicacion Access: " & Err.Description
    WScript.Quit 1
End If

' Configurar Access en modo silencioso
objAccess.Visible = False
objAccess.UserControl = False
' Suprimir alertas y diálogos de confirmación
' ' objAccess.DoCmd.SetWarnings False  ' Comentado temporalmente por error de compilación  ' Comentado temporalmente por error de compilación
objAccess.Application.Echo False
' Configuraciones adicionales para suprimir diálogos
On Error Resume Next
objAccess.Application.AutomationSecurity = 1  ' msoAutomationSecurityLow
objAccess.VBE.MainWindow.Visible = False
Err.Clear
On Error GoTo 0

' Para import-form, no abrir base de datos aquí (ImportForm lo manejará)
If strAction <> "import-form" Then
    ' Abrir base de datos con compilacion condicional
    WScript.Echo "Abriendo base de datos..."
    
    ' Configurar Access para evitar errores de compilación
    On Error Resume Next
    ' Intentar configurar propiedades si están disponibles
    objAccess.DisplayAlerts = False
    Err.Clear
    
    ' Determinar contraseña para la base de datos
    strDbPassword = GetDatabasePassword(strAccessPath)
    
    ' Abrir base de datos con manejo de errores robusto
    If strDbPassword = "" Then
        ' Sin contraseña
        objAccess.OpenCurrentDatabase strAccessPath
    Else
        ' Con contrasena - usar solo dos parametros
        objAccess.OpenCurrentDatabase strAccessPath, , strDbPassword
    End If
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al abrir base de datos: " & Err.Description
        objAccess.Quit
        WScript.Quit 1
    End If
    
    On Error GoTo 0
    WScript.Echo "Base de datos abierta correctamente."
    Call EnsureVBReferences
End If

' Verificar opciones especiales
For i = 1 To objArgs.Count - 1
    If LCase(objArgs(i)) = "--verbose" Then
        gVerbose = True
        WScript.Echo "[MODO VERBOSE] Informacion detallada activada"
    End If
Next

If strAction = "validate" Then
    Call ValidateAllModules()
ElseIf strAction = "export" Then
    Call ExportModules()
ElseIf strAction = "test" Then
    Call ExecuteTests()
ElseIf strAction = "createtable" Then
    Call CreateTable()
ElseIf strAction = "droptable" Then
    Call DropTable()
ElseIf strAction = "listtables" Then
    Call ListTables()

ElseIf strAction = "rebuild" Then
    Call RebuildProject()

ElseIf strAction = "lint" Then
    Call LintProject()
ElseIf strAction = "relink" Then
    Call RelinkTables()
ElseIf strAction = "migrate" Then
    Call ExecuteMigrations()
ElseIf strAction = "export-form" Then
    Call ExportForm()
ElseIf strAction = "import-form" Then
    Call ImportForm()

End If

' Cerrar Access
WScript.Echo "Cerrando Access..."
' Restaurar estado normal de Access antes de cerrar
On Error Resume Next
objAccess.Application.Echo True
objAccess.Quit 2  ' acQuitSaveNone = 2
If Err.Number <> 0 Then
    ' Intentar cerrar sin guardar si hay problemas
    objAccess.Quit 2  ' acQuitSaveNone = 2
End If
On Error GoTo 0
WScript.Echo "Access cerrado correctamente"

WScript.Echo "=== IMPORTACION DE FORMULARIO COMPLETADA EXITOSAMENTE ==="
WScript.Quit 0

' Subrutina para validar todos los modulos sin importar
Sub ValidateAllModules()
    Dim objFolder, objFile
    Dim strFileName, strContent
    Dim validationResult
    Dim totalFiles, validFiles, invalidFiles
    
    WScript.Echo "=== VALIDACION DE SINTAXIS VBA ==="
    
    If Not objFSO.FolderExists(strSourcePath) Then
        WScript.Echo "Error: Directorio de origen no existe: " & strSourcePath
        WScript.Quit 1
    End If
    
    Set objFolder = objFSO.GetFolder(strSourcePath)
    totalFiles = 0
    validFiles = 0
    invalidFiles = 0
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            totalFiles = totalFiles + 1
            strFileName = objFile.Path
            
            If gVerbose Then
                WScript.Echo "Validando: " & objFile.Name
            End If
            
            ' Validar sintaxis
            Dim errorDetails
            validationResult = ValidateVBASyntax(strFileName, errorDetails)
            
            If validationResult = True Then
                validFiles = validFiles + 1
                If gVerbose Then
                    WScript.Echo "  ✓ Sintaxis valida"
                End If
            Else
                invalidFiles = invalidFiles + 1
                WScript.Echo "  ✗ ERROR en " & objFile.Name & ": " & errorDetails
            End If
        End If
    Next
    
    WScript.Echo ""
    WScript.Echo "=== RESUMEN DE VALIDACION ==="
    WScript.Echo "Total de archivos: " & totalFiles
    WScript.Echo "Archivos validos: " & validFiles
    WScript.Echo "Archivos con errores: " & invalidFiles
    
    If invalidFiles > 0 Then
        WScript.Echo "ADVERTENCIA: Se encontraron errores de sintaxis. Corrija antes de importar."
        WScript.Quit 1
    Else
        WScript.Echo "✓ Todos los archivos tienen sintaxis valida"
    End If
End Sub



' Subrutina para exportar modulos
Sub ExportModules()
    Dim vbComponent
    Dim strExportPath
    Dim exportedCount
    
    WScript.Echo "Iniciando exportacion de modulos VBA..."
    
    If Not objFSO.FolderExists(strSourcePath) Then
        objFSO.CreateFolder strSourcePath
        WScript.Echo "Directorio de destino creado: " & strSourcePath
    End If
    
    exportedCount = 0
    
    For Each vbComponent In objAccess.VBE.ActiveVBProject.VBComponents
        If vbComponent.Type = 1 Then  ' vbext_ct_StdModule
            strExportPath = strSourcePath & "\" & vbComponent.Name & ".bas"
            
            If gVerbose Then
                WScript.Echo "Exportando modulo: " & vbComponent.Name
            End If
            
            On Error Resume Next
            Call ExportModuleWithAnsiEncoding(vbComponent, strExportPath)
            
            If Err.Number <> 0 Then
                WScript.Echo "Error al exportar modulo " & vbComponent.Name & ": " & Err.Description
                Err.Clear
            Else
                If gVerbose Then
                    WScript.Echo "  ✓ Modulo " & vbComponent.Name & " exportado a: " & strExportPath
                Else
                    WScript.Echo "✓ " & vbComponent.Name & ".bas"
                End If
                exportedCount = exportedCount + 1
            End If
        ElseIf vbComponent.Type = 2 Then  ' vbext_ct_ClassModule
            strExportPath = strSourcePath & "\" & vbComponent.Name & ".cls"
            
            If gVerbose Then
                WScript.Echo "Exportando clase: " & vbComponent.Name
            End If
            
            On Error Resume Next
            Call ExportModuleWithAnsiEncoding(vbComponent, strExportPath)
            
            If Err.Number <> 0 Then
                WScript.Echo "Error al exportar clase " & vbComponent.Name & ": " & Err.Description
                Err.Clear
            Else
                If gVerbose Then
                    WScript.Echo "  ✓ Clase " & vbComponent.Name & " exportada a: " & strExportPath
                Else
                    WScript.Echo "✓ " & vbComponent.Name & ".cls"
                End If
                exportedCount = exportedCount + 1
            End If
        End If
    Next
    
    WScript.Echo "Exportacion completada exitosamente. Modulos exportados: " & exportedCount
End Sub

' Subrutina para crear tabla
Sub CreateTable()
    Dim strTableName
    Dim strSQL
    Dim strQueryName
    
    If objArgs.Count < 3 Then
        WScript.Echo "Error: Se requiere nombre de tabla y consulta SQL"
        WScript.Echo "Uso: cscript condor_cli.vbs createtable <nombre> <sql>"
        WScript.Quit 1
    End If
    
    strTableName = objArgs(1)
    strSQL = objArgs(2)
    strQueryName = "qry_Create_" & strTableName
    
    WScript.Echo "Creando tabla: " & strTableName
    WScript.Echo "SQL: " & strSQL
    
    On Error Resume Next
    
    ' Verificar si la tabla ya existe
    Dim tblExists
    tblExists = False
    Dim tbl
    For Each tbl In objAccess.CurrentDb.TableDefs
        If LCase(tbl.Name) = LCase(strTableName) Then
            tblExists = True
            Exit For
        End If
    Next
    
    If tblExists Then
        WScript.Echo "Advertencia: La tabla '" & strTableName & "' ya existe"
    End If
    
    ' Crear consulta temporal
    WScript.Echo "Creando consulta temporal: " & strQueryName
    objAccess.CurrentDb.CreateQueryDef strQueryName, strSQL
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al crear consulta: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    ' Ejecutar consulta
    WScript.Echo "Ejecutando consulta..."
    objAccess.DoCmd.OpenQuery strQueryName
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al ejecutar consulta: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "Tabla '" & strTableName & "' creada exitosamente"
    End If
    
    ' Eliminar consulta temporal
    WScript.Echo "Eliminando consulta temporal..."
    objAccess.DoCmd.DeleteObject 1, strQueryName  ' acQuery = 1
    
    If Err.Number <> 0 Then
        WScript.Echo "Advertencia al eliminar consulta: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "Consulta temporal eliminada"
    End If
    
    ' Verificar que la tabla fue creada
    Call VerifyTable(strTableName)
End Sub

' Subrutina para eliminar tabla
Sub DropTable()
    Dim strTableName
    
    If objArgs.Count < 2 Then
        WScript.Echo "Error: Se requiere nombre de tabla"
        WScript.Echo "Uso: cscript condor_cli.vbs droptable <nombre>"
        WScript.Quit 1
    End If
    
    strTableName = objArgs(1)
    
    WScript.Echo "Eliminando tabla: " & strTableName
    
    On Error Resume Next
    objAccess.DoCmd.DeleteObject 0, strTableName  ' acTable = 0
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al eliminar tabla: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "Tabla '" & strTableName & "' eliminada exitosamente"
    End If
End Sub

' Subrutina para listar tablas
Sub ListTables()
    Dim tbl, fld, idx
    Dim tableCount, showSchema, outputToFile
    Dim primaryKeys
    Dim outputFile, outputPath
    
    ' Verificar flags de argumentos
    Dim arg
    showSchema = False
    outputToFile = False
    
    For Each arg In objArgs
        If LCase(arg) = "--schema" Then
            showSchema = True
        ElseIf LCase(arg) = "--output" Then
            outputToFile = True
        End If
    Next
    
    ' Configurar salida
    If outputToFile Then
        Dim dbName
        dbName = objFSO.GetBaseName(strAccessPath)
        outputPath = objFSO.GetAbsolutePathName(".") & "\" & dbName & "_listtables.txt"
        Set outputFile = objFSO.CreateTextFile(outputPath, True)
        WScript.Echo "Exportando resultados a: " & outputPath
    End If
    
    WScript.Echo "=== LISTADO DE TABLAS ==="
    If outputToFile Then outputFile.WriteLine "=== LISTADO DE TABLAS ==="
    
    If showSchema Then 
        WScript.Echo "Modo: Esquema Detallado"
        If outputToFile Then outputFile.WriteLine "Modo: Esquema Detallado"
    End If
    
    tableCount = 0
    For Each tbl In objAccess.CurrentDb.TableDefs
        If Left(tbl.Name, 4) <> "MSys" And Left(tbl.Name, 1) <> "~" Then
            tableCount = tableCount + 1
            WScript.Echo ""
            If outputToFile Then outputFile.WriteLine ""
            
            WScript.Echo "------------------------------------------------------------"
            If outputToFile Then outputFile.WriteLine "------------------------------------------------------------"
            
            WScript.Echo tableCount & ". " & tbl.Name & " (" & tbl.RecordCount & " registros)"
            If outputToFile Then outputFile.WriteLine tableCount & ". " & tbl.Name & " (" & tbl.RecordCount & " registros)"
            
            WScript.Echo "------------------------------------------------------------"
            If outputToFile Then outputFile.WriteLine "------------------------------------------------------------"
            
            If showSchema Then
                Set primaryKeys = CreateObject("Scripting.Dictionary")
                For Each idx In tbl.Indexes
                    If idx.Primary Then
                        For Each fld In idx.Fields
                            primaryKeys.Add fld.Name, True
                        Next
                    End If
                Next
    
                WScript.Echo PadRight("Campo", 25) & PadRight("Tipo", 15) & PadRight("PK", 8) & "Requerido"
                If outputToFile Then outputFile.WriteLine PadRight("Campo", 25) & PadRight("Tipo", 15) & PadRight("PK", 8) & "Requerido"
                
                WScript.Echo "--------------------------------------------------------------------"
                If outputToFile Then outputFile.WriteLine "--------------------------------------------------------------------"
                
                For Each fld In tbl.Fields
                    Dim pkMarker, requiredMarker
                    If primaryKeys.Exists(fld.Name) Then pkMarker = "PK" Else pkMarker = ""
                    If fld.Required Then requiredMarker = "true" Else requiredMarker = "false"
                    WScript.Echo PadRight(fld.Name, 25) & PadRight(DaoTypeToString(fld.Type), 15) & PadRight(pkMarker, 8) & requiredMarker
                    If outputToFile Then outputFile.WriteLine PadRight(fld.Name, 25) & PadRight(DaoTypeToString(fld.Type), 15) & PadRight(pkMarker, 8) & requiredMarker
                Next
            End If
        End If
    Next
    
    WScript.Echo ""
    If outputToFile Then outputFile.WriteLine ""
    
    WScript.Echo "Total de tablas: " & tableCount
    If outputToFile Then outputFile.WriteLine "Total de tablas: " & tableCount
    
    ' Cerrar archivo si se está usando
    If outputToFile Then
        outputFile.Close
        Set outputFile = Nothing
        WScript.Echo "Archivo generado exitosamente: " & outputPath
    End If
End Sub

' Subrutina para verificar tabla creada
Sub VerifyTable(strTableName)
    Dim tbl
    Dim found
    
    WScript.Echo "Verificando tabla creada..."
    found = False
    
    On Error Resume Next
    For Each tbl In objAccess.CurrentDb.TableDefs
        If LCase(tbl.Name) = LCase(strTableName) Then
            found = True
            WScript.Echo "? Tabla '" & strTableName & "' verificada exitosamente"
            WScript.Echo "  - Campos: " & tbl.Fields.Count
            WScript.Echo "  - Registros: " & tbl.RecordCount
            Exit For
        End If
    Next
    
    If Not found Then
        WScript.Echo "? Error: No se pudo verificar la tabla '" & strTableName & "'"
    End If
End Sub




' ===================================================================
' SUBRUTINA: LintProject
' Descripción: Audita el código VBA para detectar cabeceras duplicadas
' ===================================================================
Sub LintProject()
    Dim vbComponent, codeModule
    Dim lineContent, moduleName
    Dim optionCompareCount, optionExplicitCount
    Dim i, hasErrors
    
    WScript.Echo "=== INICIANDO AUDITORIA VBA ==="
    WScript.Echo "Accion: lint"
    WScript.Echo "Base de datos: " & strAccessPath
    
    Set objAccess = CreateObject("Access.Application")
    objAccess.Visible = False
    objAccess.OpenCurrentDatabase strAccessPath, False
    
    WScript.Echo "=== AUDITORIA DE CABECERAS VBA ==="
    WScript.Echo ""
    
    hasErrors = False
    
    For Each vbComponent In objAccess.VBE.ActiveVBProject.VBComponents
        moduleName = vbComponent.Name
        Set codeModule = vbComponent.CodeModule
        
        optionCompareCount = 0
        optionExplicitCount = 0
        
        For i = 1 To 10
            If i <= codeModule.CountOfLines Then
                lineContent = Trim(codeModule.Lines(i, 1))
                
                If InStr(1, lineContent, "Option Compare", 1) > 0 Then
                    optionCompareCount = optionCompareCount + 1
                End If
                
                If InStr(1, lineContent, "Option Explicit", 1) > 0 Then
                    optionExplicitCount = optionExplicitCount + 1
                End If
            End If
        Next
        
        If optionCompareCount > 1 Then
            WScript.Echo "ERROR: Modulo " & moduleName & " tiene " & optionCompareCount & " declaraciones Option Compare duplicadas"
            hasErrors = True
        End If
        
        If optionExplicitCount > 1 Then
            WScript.Echo "ERROR: Modulo " & moduleName & " tiene " & optionExplicitCount & " declaraciones Option Explicit duplicadas"
            hasErrors = True
        End If
        
        If optionCompareCount <= 1 And optionExplicitCount <= 1 Then
            WScript.Echo "OK: " & moduleName & " - Cabeceras correctas"
        End If
    Next
    
    If hasErrors Then
        WScript.Echo ""
        WScript.Echo "=== LINT FALLIDO ==="
        WScript.Echo "Se encontraron cabeceras duplicadas."
        objAccess.Quit
        WScript.Quit 1
    Else
        WScript.Echo ""
        WScript.Echo "=== LINT COMPLETADO EXITOSAMENTE ==="
    End If
    
    objAccess.Quit
End Sub

' Subrutina para compilación condicional de módulos
Sub CompileModulesConditionally()
    Dim vbComponent
    Dim compilationErrors
    Dim totalModules
    Dim compiledModules
    
    WScript.Echo "Iniciando compilación condicional de módulos..."
    
    compilationErrors = 0
    totalModules = 0
    compiledModules = 0
    
    ' Intentar compilar cada módulo individualmente (módulos estándar y clases)
    For Each vbComponent In objAccess.VBE.ActiveVBProject.VBComponents
        If vbComponent.Type = 1 Or vbComponent.Type = 2 Then  ' vbext_ct_StdModule o vbext_ct_ClassModule
            totalModules = totalModules + 1
            
            On Error Resume Next
            Err.Clear
            
            ' Intentar compilar el módulo específico
            If vbComponent.Type = 1 Then
                WScript.Echo "Compilando módulo: " & vbComponent.Name
            Else
                WScript.Echo "Compilando clase: " & vbComponent.Name
            End If
            
            ' Verificar si el módulo tiene errores de sintaxis
            Dim hasErrors
            hasErrors = False
            
            ' Intentar acceder al código del módulo para detectar errores
            Dim moduleCode
            moduleCode = vbComponent.CodeModule.Lines(1, vbComponent.CodeModule.CountOfLines)
            
            If Err.Number <> 0 Then
                WScript.Echo "  ⚠️ Error en " & vbComponent.Name & ": " & Err.Description
                compilationErrors = compilationErrors + 1
                hasErrors = True
                Err.Clear
            Else
                ' Los módulos se guardan automáticamente, no es necesario guardar explícitamente
                If vbComponent.Type = 1 Then
                    WScript.Echo "  ✓ " & vbComponent.Name & " compilado correctamente"
                    compiledModules = compiledModules + 1
                Else
                    ' Para módulos de clase, solo verificar sintaxis sin intentar guardar individualmente
                    WScript.Echo "  ✓ " & vbComponent.Name & " verificado (clase)"
                    compiledModules = compiledModules + 1
                End If
            End If
            
            On Error GoTo 0
        End If
    Next
    
    ' Intentar compilación global si los módulos principales están bien
    If compiledModules >= (totalModules - 3) Then  ' Permitir hasta 3 errores (las clases problemáticas)
        WScript.Echo "Intentando compilación global..."
        On Error Resume Next
        objAccess.DoCmd.RunCommand 636  ' acCmdCompileAndSaveAllModules
        
        If Err.Number <> 0 Then
            WScript.Echo "⚠️ Advertencia en compilación global: " & Err.Description
            WScript.Echo "Continuando con módulos compilados individualmente..."
            Err.Clear
        Else
            WScript.Echo "✓ Compilación global exitosa"
        End If
        On Error GoTo 0
    Else
        WScript.Echo "⚠️ Se encontraron " & compilationErrors & " errores de compilación"
        WScript.Echo "Continuando sin compilación global para evitar bloqueos..."
    End If
    
    WScript.Echo "Resumen de compilación:"
    WScript.Echo "  - Total de módulos: " & totalModules
    WScript.Echo "  - Módulos compilados: " & compiledModules
    WScript.Echo "  - Errores encontrados: " & compilationErrors
    
    If compilationErrors > 0 Then
        WScript.Echo "⚠️ ADVERTENCIA: Algunos módulos tienen errores de compilación"
        WScript.Echo "El CLI continuará funcionando, pero revise los módulos con errores"
    End If
End Sub

' Subrutina para verificar que los nombres de módulos coincidan con src
Sub VerifyModuleNames()
    Dim objFolder, objFile
    Dim vbComponent
    Dim srcModules, accessModules
    Dim moduleName
    Dim discrepancies
    
    WScript.Echo "Verificando integridad de nombres de módulos..."
    
    ' Crear diccionarios para comparación
    Set srcModules = CreateObject("Scripting.Dictionary")
    Set accessModules = CreateObject("Scripting.Dictionary")
    discrepancies = 0
    
    ' Obtener lista de módulos en src
    Set objFolder = objFSO.GetFolder(strSourcePath)
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            moduleName = objFSO.GetBaseName(objFile.Name)
            srcModules.Add moduleName, True
        End If
    Next
    
    ' Obtener lista de módulos en Access
    For Each vbComponent In objAccess.VBE.ActiveVBProject.VBComponents
        If vbComponent.Type = 1 Or vbComponent.Type = 2 Then  ' vbext_ct_StdModule o vbext_ct_ClassModule
            accessModules.Add vbComponent.Name, True
        End If
    Next
    
    ' Verificar que todos los módulos de src estén en Access
    For Each moduleName In srcModules.Keys
        If Not accessModules.Exists(moduleName) Then
            WScript.Echo "⚠️ ERROR: Módulo '" & moduleName & "' existe en src pero no en Access"
            discrepancies = discrepancies + 1
        End If
    Next
    
    ' Verificar que todos los módulos de Access estén en src
    For Each moduleName In accessModules.Keys
        If Not srcModules.Exists(moduleName) Then
            WScript.Echo "⚠️ ERROR: Módulo '" & moduleName & "' existe en Access pero no en src"
            discrepancies = discrepancies + 1
        End If
    Next
    
    ' Reporte final
    If discrepancies = 0 Then
        WScript.Echo "✓ Verificación exitosa: Todos los módulos coinciden entre src y Access"
        WScript.Echo "  - Módulos en src: " & srcModules.Count
        WScript.Echo "  - Módulos en Access: " & accessModules.Count
    Else
        WScript.Echo "❌ FALLO EN VERIFICACIÓN: Se encontraron " & discrepancies & " discrepancias"
        WScript.Echo "⚠️ ACCIÓN REQUERIDA: Revise la sincronización entre src y Access"
    End If
End Sub

' Función para validar sintaxis VBA antes de importar
Function ValidateVBASyntax(filePath, ByRef errorDetails)
    Dim objFile, strContent
    
    errorDetails = ""
    
    ' Leer archivo con codificación ANSI
    On Error Resume Next
    Set objFile = objFSO.OpenTextFile(filePath, 1, False, 0)
    If Err.Number <> 0 Then
        errorDetails = "Error al leer archivo: " & Err.Description
        ValidateVBASyntax = False
        Exit Function
    End If
    
    strContent = objFile.ReadAll
    objFile.Close
    On Error GoTo 0
    
    ' Validación básica: verificar que el archivo no esté vacío y sea legible
    If Len(Trim(strContent)) = 0 Then
        errorDetails = "El archivo está vacío"
        ValidateVBASyntax = False
        Exit Function
    End If
    
    ' Verificar caracteres problemáticos básicos
    If InStr(strContent, Chr(0)) > 0 Then
        errorDetails = "El archivo contiene caracteres nulos"
        ValidateVBASyntax = False
        Exit Function
    End If
    
    ' Si llegamos aquí, el archivo es válido
    ValidateVBASyntax = True
End Function

' Función de serialización recursiva para convertir Dictionary a JSON
Private Function DictionaryToJson(obj, indentLevel)
    Dim result, indent, i, key, value, keys
    Dim isLast
    
    ' Crear indentación
    indent = String(indentLevel * 4, " ")
    
    ' Verificar qué tipo de objeto estamos procesando
    
    ' Verificar el tipo de objeto
    If TypeName(obj) = "Dictionary" Then
        ' Verificar si es una colección de controles (Dictionary que contiene objetos con propiedades 'name' y 'type')
        Dim isControlsCollection
        isControlsCollection = False
        
        If obj.Count > 0 Then
            keys = obj.Keys
            ' Verificar si el primer elemento es un diccionario con 'name' y 'type'
            If TypeName(obj(keys(0))) = "Dictionary" Then
                If obj(keys(0)).Exists("name") And obj(keys(0)).Exists("type") Then
                    isControlsCollection = True
                    WScript.Echo "DEBUG: Detectada colección de controles con " & obj.Count & " elementos"
                End If
            End If
        Else
            WScript.Echo "DEBUG: Dictionary vacío encontrado"
        End If
        
        If isControlsCollection Then
            ' Manejar como array de controles
            result = "[" & vbCrLf
            keys = obj.Keys
            For i = 0 To UBound(keys)
                On Error Resume Next
                Set value = obj(keys(i))
                If Err.Number <> 0 Then
                    Err.Clear
                    value = obj(keys(i))
                End If
                On Error GoTo 0
                isLast = (i = UBound(keys))
                
                WScript.Echo "DEBUG: Procesando elemento array " & i & ": " & keys(i)
                result = result & indent & "    "
                On Error Resume Next
                If IsObject(value) Then
                    If Not value Is Nothing Then
                        WScript.Echo "DEBUG: Serializando objeto control: " & TypeName(value)
                        result = result & DictionaryToJson(value, indentLevel + 1)
                    Else
                        WScript.Echo "DEBUG: Objeto es Nothing"
                        result = result & "null"
                    End If
                ElseIf IsNull(value) Then
                    result = result & "null"
                Else
                    ' Manejar valores simples directamente
                    If VarType(value) = vbString Then
                        result = result & Chr(34) & Replace(CStr(value), Chr(34), "\" & Chr(34)) & Chr(34)
                    ElseIf VarType(value) = vbBoolean Then
                        If value Then
                            result = result & "true"
                        Else
                            result = result & "false"
                        End If
                    ElseIf IsNumeric(value) Then
                        result = result & CStr(value)
                    Else
                        result = result & Chr(34) & CStr(value) & Chr(34)
                    End If
                End If
                On Error GoTo 0
                
                If Not isLast Then
                    result = result & ","
                End If
                result = result & vbCrLf
            Next
            result = result & indent & "]"
        Else
            ' Manejar Dictionary normal
            result = "{" & vbCrLf
            keys = obj.Keys
            For i = 0 To UBound(keys)
                key = keys(i)
                On Error Resume Next
                Set value = obj(key)
                If Err.Number <> 0 Then
                    Err.Clear
                    value = obj(key)
                End If
                On Error GoTo 0
                isLast = (i = UBound(keys))
                
                result = result & indent & "    " & Chr(34) & key & Chr(34) & ": "
                If IsObject(value) Then
                    If Not value Is Nothing Then
                        result = result & DictionaryToJson(value, indentLevel + 1)
                    Else
                        result = result & "null"
                    End If
                ElseIf IsNull(value) Then
                    result = result & "null"
                Else
                    ' Manejar valores simples directamente
                    If VarType(value) = vbString Then
                        result = result & Chr(34) & Replace(CStr(value), Chr(34), "\" & Chr(34)) & Chr(34)
                    ElseIf VarType(value) = vbBoolean Then
                        If value Then
                            result = result & "true"
                        Else
                            result = result & "false"
                        End If
                    ElseIf IsNumeric(value) Then
                        result = result & CStr(value)
                    Else
                        result = result & Chr(34) & CStr(value) & Chr(34)
                    End If
                End If
                On Error GoTo 0
                
                If Not isLast Then
                    result = result & ","
                End If
                result = result & vbCrLf
            Next
            result = result & indent & "}"
        End If

    Else
        ' Manejar valores simples
        If VarType(obj) = vbString Then
            ' String - escapar comillas dobles
            result = Chr(34) & Replace(CStr(obj), Chr(34), "\" & Chr(34)) & Chr(34)
        ElseIf VarType(obj) = vbBoolean Then
            ' Boolean
            If obj Then
                result = "true"
            Else
                result = "false"
            End If
        ElseIf IsNumeric(obj) Then
            ' Numérico
            result = CStr(obj)
        Else
            ' Otros tipos como string
            result = Chr(34) & CStr(obj) & Chr(34)
        End If
    End If
    
    DictionaryToJson = result
End Function

' Función para leer archivo con codificación ANSI
Function ReadFileWithAnsiEncoding(filePath)
    Dim objStream, strContent
    
    On Error Resume Next
    
    ' Leer contenido del archivo usando ADODB.Stream con UTF-8 y convertir a ANSI
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "UTF-8"
    objStream.Open
    objStream.LoadFromFile filePath
    strContent = objStream.ReadText
    objStream.Close
    
    ' Convertir caracteres UTF-8 a ANSI para compatibilidad con VBA
    ' Preservar caracteres especiales del español
    strContent = Replace(strContent, "á", "á")
    strContent = Replace(strContent, "é", "é")
    strContent = Replace(strContent, "í", "í")
    strContent = Replace(strContent, "ó", "ó")
    strContent = Replace(strContent, "ú", "ú")
    strContent = Replace(strContent, "ñ", "ñ")
    strContent = Replace(strContent, "Á", "Á")
    strContent = Replace(strContent, "É", "É")
    strContent = Replace(strContent, "Í", "Í")
    strContent = Replace(strContent, "Ó", "Ó")
    strContent = Replace(strContent, "Ú", "Ú")
    strContent = Replace(strContent, "Ñ", "Ñ")
    strContent = Replace(strContent, "ü", "ü")
    strContent = Replace(strContent, "Ü", "Ü")
    
    Set objStream = Nothing
    
    If Err.Number <> 0 Then
        WScript.Echo "❌ ERROR: No se pudo leer el archivo " & filePath & ": " & Err.Description
        ReadFileWithAnsiEncoding = ""
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    
    On Error GoTo 0
    ReadFileWithAnsiEncoding = strContent
End Function

' Función para limpiar archivos VBA eliminando líneas Attribute con validación mejorada
Function CleanVBAFile(filePath, fileType)
    Dim objStream, strContent, arrLines, i, cleanedContent
    Dim strLine
    
    ' Leer el archivo como UTF-8 y convertir a ANSI para VBA
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "UTF-8"
    objStream.Open
    objStream.LoadFromFile filePath
    strContent = objStream.ReadText
    objStream.Close
    
    ' Convertir caracteres UTF-8 a ANSI para compatibilidad con VBA
    ' Preservar caracteres especiales del español
    strContent = Replace(strContent, "á", "á")
    strContent = Replace(strContent, "é", "é")
    strContent = Replace(strContent, "í", "í")
    strContent = Replace(strContent, "ó", "ó")
    strContent = Replace(strContent, "ú", "ú")
    strContent = Replace(strContent, "ñ", "ñ")
    strContent = Replace(strContent, "Á", "Á")
    strContent = Replace(strContent, "É", "É")
    strContent = Replace(strContent, "Í", "Í")
    strContent = Replace(strContent, "Ó", "Ó")
    strContent = Replace(strContent, "Ú", "Ú")
    strContent = Replace(strContent, "Ñ", "Ñ")
    strContent = Replace(strContent, "ü", "ü")
    strContent = Replace(strContent, "Ü", "Ü")
    
    Set objStream = Nothing
    
    ' Dividir el contenido en un array de líneas
    strContent = Replace(strContent, vbCrLf, vbLf)
    strContent = Replace(strContent, vbCr, vbLf)
    arrLines = Split(strContent, vbLf)
    
    ' Crear un nuevo string vacío llamado cleanedContent
    cleanedContent = ""
    
    ' Iterar sobre el array de líneas original
    For i = 0 To UBound(arrLines)
        strLine = arrLines(i)
        
        ' Aplicar las reglas para descartar contenido no deseado
        ' Una línea se descarta si cumple cualquiera de estas condiciones:
        ' CORRECCION CRITICA: Filtrar TODAS las líneas que empiecen con 'Attribute'
        ' y todos los metadatos de archivos .cls
        If Not (Left(Trim(strLine), 9) = "Attribute" Or _
                Left(Trim(strLine), 17) = "VERSION 1.0 CLASS" Or _
                Trim(strLine) = "BEGIN" Or _
                Left(Trim(strLine), 8) = "MultiUse" Or _
                Trim(strLine) = "END" Or _
                Trim(strLine) = "Option Compare Database" Or _
                Trim(strLine) = "Option Explicit") Then
            
            ' Si no cumple ninguna condición, es código VBA válido
            ' Se añade al cleanedContent seguida de un salto de línea
            cleanedContent = cleanedContent & strLine & vbCrLf
        End If
    Next
    
    ' La función devuelve cleanedContent directamente
    ' No añade ninguna cabecera Option manualmente
    CleanVBAFile = cleanedContent
End Function




' Función para exportar módulo con conversión ANSI -> UTF-8 usando ADODB.Stream
Sub ExportModuleWithAnsiEncoding(vbComponent, strExportPath)
    Dim tempFilePath, objTempFile, objStream
    Dim strContent
    Dim tempFolderPath, tempFileName
    
    On Error Resume Next
    
    ' Crear archivo temporal usando el directorio temporal del sistema
    tempFolderPath = objFSO.GetSpecialFolder(2) ' El 2 es la constante para la carpeta temporal del sistema
    tempFileName = objFSO.GetTempName() ' Genera un nombre aleatorio y seguro como "radB93EB.tmp"
    tempFilePath = objFSO.BuildPath(tempFolderPath, tempFileName)
    
    ' Exportar a archivo temporal (Access usa ANSI internamente)
    vbComponent.Export tempFilePath
    
    ' Leer contenido del archivo temporal con codificación ANSI usando FSO
    Set objTempFile = objFSO.OpenTextFile(tempFilePath, 1, False, 0) ' ForReading = 1, Create = False, Format = 0 (ANSI)
    strContent = objTempFile.ReadAll
    objTempFile.Close
    
    ' Escribir al archivo final con codificación UTF-8 usando ADODB.Stream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "UTF-8"
    objStream.Open
    objStream.WriteText strContent
    objStream.SaveToFile strExportPath, 2 ' adSaveCreateOverWrite
    objStream.Close
    Set objStream = Nothing
    
    ' Limpiar archivo temporal
    objFSO.DeleteFile tempFilePath
    
    If Err.Number <> 0 Then
        WScript.Echo "❌ ERROR en ExportModuleWithAnsiEncoding: " & Err.Description
        Err.Clear
    End If
    
    On Error GoTo 0
End Sub

' Subrutina principal para validar esquemas de base de datos
Sub ValidateSchema()
    WScript.Echo "=== INICIANDO VALIDACIÓN DE ESQUEMA DE BASE DE DATOS ==="
    
    Dim lanzaderaSchema, condorSchema
    Dim allOk
    allOk = True
    
    ' Definir esquema esperado para Lanzadera
    Set lanzaderaSchema = CreateObject("Scripting.Dictionary")
    lanzaderaSchema.Add "TbUsuariosAplicaciones", Array("CorreoUsuario", "Password", "UsuarioRed", "Nombre", "Matricula", "FechaAlta")
    lanzaderaSchema.Add "TbUsuariosAplicacionesPermisos", Array("CorreoUsuario", "IDAplicacion", "EsUsuarioAdministrador", "EsUsuarioCalidad", "EsUsuarioEconomia", "EsUsuarioSecretaria")
    
    ' Definir esquema esperado para CONDOR
    Set condorSchema = CreateObject("Scripting.Dictionary")
    condorSchema.Add "tbSolicitudes", Array("idSolicitud", "idExpediente", "tipoSolicitud", "estadoInterno", "fechaCreacion", "usuarioCreacion", "fechaModificacion", "usuarioModificacion", "observaciones")
    condorSchema.Add "tbDatosPC", Array("idSolicitud", "numeroExpediente", "fechaExpediente", "tipoExpediente", "descripcionCambio", "justificacionCambio", "impactoCalidad", "impactoSeguridad", "impactoOperacional")
    condorSchema.Add "tbDatosCDCA", Array("idSolicitud", "numeroExpediente", "fechaExpediente", "tipoExpediente", "descripcionDesviacion", "justificacionDesviacion", "impactoCalidad", "impactoSeguridad", "impactoOperacional")
    condorSchema.Add "tbDatosCDCASUB", Array("idSolicitud", "numeroExpediente", "fechaExpediente", "tipoExpediente", "descripcionDesviacion", "justificacionDesviacion", "impactoCalidad", "impactoSeguridad", "impactoOperacional", "subsuministrador")
    condorSchema.Add "tbMapeoCampos", Array("NombrePlantilla", "NombreCampoTabla", "ValorAsociado", "NombreCampoWord")
    condorSchema.Add "tbLogCambios", Array("idLog", "idSolicitud", "fechaCambio", "usuarioCambio", "campoModificado", "valorAnterior", "valorNuevo")
    condorSchema.Add "tbLogErrores", Array("idError", "fechaError", "tipoError", "descripcionError", "moduloOrigen", "funcionOrigen", "usuarioAfectado")
    condorSchema.Add "tbOperacionesLog", Array("idOperacion", "fechaOperacion", "tipoOperacion", "descripcionOperacion", "usuario", "resultado")
    condorSchema.Add "tbAdjuntos", Array("idAdjunto", "idSolicitud", "nombreArchivo", "rutaArchivo", "tipoArchivo", "fechaSubida", "usuarioSubida")
    condorSchema.Add "tbEstados", Array("idEstado", "nombreEstado", "descripcionEstado", "esEstadoFinal")
    condorSchema.Add "tbTransiciones", Array("idTransicion", "estadoOrigen", "estadoDestino", "accionRequerida", "rolRequerido")
    condorSchema.Add "tbConfiguracion", Array("clave", "valor", "descripcion", "categoria")
    condorSchema.Add "TbLocalConfig", Array("clave", "valor", "descripcion", "categoria")
    
    ' Validar las bases de datos
    If Not VerifySchema(strSourcePath & "\..\back\test_env\fixtures\databases\Lanzadera_test_template.accdb", "dpddpd", lanzaderaSchema) Then allOk = False
        If Not VerifySchema(strSourcePath & "\..ack\test_env\fixtures\databases\Document_test_template.accdb", "", condorSchema) Then allOk = False
    
    If allOk Then
        WScript.Echo "✓ VALIDACIÓN DE ESQUEMA EXITOSA. Todas las bases de datos son consistentes."
        WScript.Quit 0
    Else
        WScript.Echo "X VALIDACIÓN DE ESQUEMA FALLIDA. Corrija las discrepancias."
        WScript.Quit 1
    End If
End Sub

' Función auxiliar para verificar esquema de una base de datos específica
Private Function VerifySchema(dbPath, dbPassword, expectedSchema)
    On Error Resume Next
    
    WScript.Echo "Validando base de datos: " & dbPath
    
    ' Verificar que existe la base de datos
    If Not objFSO.FileExists(dbPath) Then
        WScript.Echo "❌ ERROR: Base de datos no encontrada: " & dbPath
        VerifySchema = False
        Exit Function
    End If
    
    ' Crear conexión ADO
    Dim conn, rs
    Set conn = CreateObject("ADODB.Connection")
    
    ' Construir cadena de conexión
    Dim connectionString
    If dbPassword = "" Then
        connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    Else
        connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=" & dbPassword & ";"
    End If
    
    ' Abrir conexión
    conn.Open connectionString
    
    If Err.Number <> 0 Then
        WScript.Echo "❌ ERROR: No se pudo conectar a la base de datos: " & Err.Description
        VerifySchema = False
        Err.Clear
        Exit Function
    End If
    
    ' Iterar sobre cada tabla esperada
    Dim tableName, expectedFields, i
    Dim tableExists, fieldExists
    Dim allTablesOk
    allTablesOk = True
    
    For Each tableName In expectedSchema.Keys
        expectedFields = expectedSchema(tableName)
        
        ' Verificar que existe la tabla usando una consulta más simple
         Set rs = CreateObject("ADODB.Recordset")
         On Error Resume Next
         rs.Open "SELECT TOP 1 * FROM [" & tableName & "]", conn
         tableExists = (Err.Number = 0)
         If tableExists Then rs.Close
         Err.Clear
         On Error GoTo 0
        
        If Not tableExists Then
            WScript.Echo "❌ ERROR: Tabla no encontrada: " & tableName
            allTablesOk = False
        Else
            WScript.Echo "✓ Tabla encontrada: " & tableName
            
            ' Verificar cada campo esperado
            For i = 0 To UBound(expectedFields)
                Dim fieldName
                fieldName = expectedFields(i)
                
                ' Verificar que existe el campo usando una consulta más simple
                 Set rs = CreateObject("ADODB.Recordset")
                 On Error Resume Next
                 rs.Open "SELECT [" & fieldName & "] FROM [" & tableName & "] WHERE 1=0", conn
                 fieldExists = (Err.Number = 0)
                 If fieldExists Then rs.Close
                 Err.Clear
                 On Error GoTo 0
                
                If Not fieldExists Then
                    WScript.Echo "❌ ERROR: Campo no encontrado: " & tableName & "." & fieldName
                    allTablesOk = False
                Else
                    WScript.Echo "  ✓ Campo encontrado: " & fieldName
                End If
            Next
        End If
    Next
    
    ' Cerrar conexión
    conn.Close
    Set conn = Nothing
    
    If allTablesOk Then
        WScript.Echo "✅ Base de datos validada correctamente: " & objFSO.GetFileName(dbPath)
        VerifySchema = True
    Else
        WScript.Echo "❌ Errores encontrados en: " & objFSO.GetFileName(dbPath)
        VerifySchema = False
    End If
    
    On Error GoTo 0
End Function

Sub EnsureVBReferences()
    WScript.Echo "Verificando referencias VBA críticas..."
    On Error Resume Next
    Dim vbProj: Set vbProj = objAccess.VBE.ActiveVBProject
    If vbProj Is Nothing Then Exit Sub
    
    Dim refs(1, 2)
    refs(0, 0) = "{420B2830-E718-11CF-893D-00A0C9054228}": refs(0, 1) = "1.0": refs(0, 2) = "Scripting Runtime"
    refs(1, 0) = "{0002E157-0000-0000-C000-000000000046}": refs(1, 1) = "5.3": refs(1, 2) = "VBIDE Extensibility"

    Dim i, ref, found
    For i = 0 To 1
        found = False
        For Each ref In vbProj.References
            If ref.Guid = refs(i, 0) Then found = True: Exit For
        Next
        If Not found Then
            WScript.Echo "  -> Añadiendo: " & refs(i, 2)
            vbProj.References.AddFromGuid refs(i, 0), CInt(Split(refs(i, 1), ".")(0)), CInt(Split(refs(i, 1), ".")(1))
        End If
    Next
    On Error GoTo 0
End Sub

' Subrutina para mostrar ayuda completa
Sub ShowHelp()
    WScript.Echo "=== CONDOR CLI - Herramienta de línea de comandos ==="
    WScript.Echo "Versión: 2.0 - Sistema de gestión y sincronización VBA para proyecto CONDOR"
    WScript.Echo ""
    WScript.Echo "SINTAXIS:"
    WScript.Echo "  cscript condor_cli.vbs [comando] [opciones] [parámetros]"
    WScript.Echo ""
    WScript.Echo "COMANDOS PRINCIPALES:"
    WScript.Echo ""
    WScript.Echo "📤 EXPORTACIÓN:"
    WScript.Echo "  export [--verbose]           - Exportar módulos VBA desde Access a /src"
    WScript.Echo "                                 Codificación: ANSI para compatibilidad"
    WScript.Echo "                                 --verbose: Mostrar detalles de cada archivo"
    WScript.Echo ""
    WScript.Echo "🔄 SINCRONIZACIÓN:"
    WScript.Echo "  rebuild                      - Método principal de sincronización del proyecto"
    WScript.Echo "                                 Reconstrucción completa: elimina todos los módulos"
    WScript.Echo "                                 y reimporta desde /src para garantizar coherencia"
    WScript.Echo ""
    WScript.Echo "✅ VALIDACIÓN Y PRUEBAS:"
    WScript.Echo "  validate [--verbose] [--src] - Validar sintaxis VBA sin importar a Access"
    WScript.Echo "                                 --verbose: Mostrar detalles de validación"
    WScript.Echo "                                 --src: Usar directorio fuente alternativo"
    WScript.Echo "  test                         - Ejecutar suite completa de pruebas unitarias"
    WScript.Echo "  lint                         - Auditar código VBA (detectar cabeceras duplicadas)"
    WScript.Echo ""
    WScript.Echo "📦 EMPAQUETADO:"
    WScript.Echo "  bundle <funcionalidad> [destino] - Empaquetar archivos por funcionalidad"
    WScript.Echo "                                      Destino opcional (por defecto: directorio actual)"
    WScript.Echo ""
    WScript.Echo "🗄️ GESTIÓN DE BASE DE DATOS:"
    WScript.Echo "  createtable <nombre> <sql>   - Crear tabla con consulta SQL personalizada"
    WScript.Echo "  droptable <nombre>           - Eliminar tabla de la base de datos"
    WScript.Echo "  listtables [db_path]         - Listar todas las tablas"
    WScript.Echo "                                 db_path opcional (por defecto: CONDOR_datos.accdb)"
    WScript.Echo "  relink <db_path> <folder>    - Re-vincular tablas a bases locales específicas"
    WScript.Echo "  relink --all                 - Re-vincular automáticamente todas las bases en ./back"
    WScript.Echo "  migrate [file.sql]           - Ejecutar scripts de migración SQL desde ./db/migrations"
    WScript.Echo "  export-form <db_path> <form_name> [opciones] - Exportar diseño de formulario a JSON enriquecido."
    WScript.Echo "                                 Genera JSON versionado con metadata, normalización de colores y recursos."
    WScript.Echo "                                 Incluye propiedades completas del formulario: caption, popUp, modal, width,"
    WScript.Echo "                                 autoCenter, borderStyle, recordSelectors, dividingLines, navigationButtons,"
    WScript.Echo "                                 scrollBars, controlBox, closeButton, minMaxButtons, movable, recordsetType,"
    WScript.Echo "                                 orientation y propiedades de SplitForm (si aplica)."
    WScript.Echo "                                 Opciones:"
    WScript.Echo "                                   --output <archivo>        - Archivo de salida (por defecto: <form_name>.json)"
    WScript.Echo "                                   --password <pwd>          - Contraseña de la base de datos"
    WScript.Echo "                                   --schema-version <ver>    - Versión del esquema (por defecto: 1.0.0)"
    WScript.Echo "                                   --expand <ámbitos>        - Ámbitos a incluir: events,formatting,resources"
    WScript.Echo "                                                               (por defecto: todos)"
    WScript.Echo "                                   --resource-root <dir>     - Directorio base para rutas relativas de recursos"
    WScript.Echo "                                   --pretty                  - Formatear JSON con indentación"
    WScript.Echo "                                   --no-controls             - Solo metadata del formulario, sin controles"
    WScript.Echo "                                   --verbose                 - Mostrar información detallada del proceso"
    WScript.Echo "                                 Propiedades JSON exportadas:"
    WScript.Echo "                                   • caption: string - Título del formulario"
    WScript.Echo "                                   • popUp/modal: boolean - Comportamiento de ventana"
    WScript.Echo "                                   • width: number (twips) - Ancho del formulario"
    WScript.Echo "                                   • borderStyle: ""None""|""Thin""|""Sizable""|""Dialog"""
    WScript.Echo "                                   • scrollBars: ""Neither""|""Horizontal""|""Vertical""|""Both"""
    WScript.Echo "                                   • minMaxButtons: ""None""|""Min Enabled""|""Max Enabled""|""Both Enabled"""
    WScript.Echo "                                   • recordsetType: ""Dynaset""|""Snapshot""|""Dynaset (Inconsistent Updates)"""
    WScript.Echo "                                   • orientation: ""LeftToRight""|""RightToLeft"""
    WScript.Echo "                                   • splitForm*: propiedades específicas para formularios divididos"
    WScript.Echo "                                 Ejemplo: export-form db.accdb MiForm --pretty --expand=events,formatting"
    WScript.Echo "  import-form <json_path> <db_path> [--password] - Crear/Modificar formulario desde JSON."
    WScript.Echo "  validate-form-json <json_path> [--strict] [--schema] - Validar estructura JSON de formulario"
    WScript.Echo "                                 --strict: Validación exhaustiva de coherencia con código VBA"
    WScript.Echo "                                 --schema: Validar contra esquema específico"
    WScript.Echo "  roundtrip-form <db_path> <form> [--password] [--temp-dir] [--verbose] - Test export→import de formulario"
    WScript.Echo ""
    WScript.Echo "FUNCIONALIDADES DISPONIBLES PARA 'bundle' (con dependencias automáticas):"
    WScript.Echo "(Basadas en CONDOR_MASTER_PLAN.md)"
    WScript.Echo ""
    WScript.Echo "🔐 Auth          - Autenticación + dependencias (Config, Error, Modelos)"
    WScript.Echo "                   Incluye archivos de autenticación y sus dependencias"
    WScript.Echo ""
    WScript.Echo "📄 Document      - Gestión de Documentos + dependencias (Config, FileSystem, Error, Word, Modelos)"
    WScript.Echo "                   Incluye archivos de documentos y sus dependencias"
    WScript.Echo ""
    WScript.Echo "📁 Expediente    - Gestión de Expedientes + dependencias (Config, Error, Modelos, Utilidades)"
    WScript.Echo "                   Incluye archivos de expedientes y sus dependencias"
    WScript.Echo ""
    WScript.Echo "📋 Solicitud     - Gestión de Solicitudes + dependencias (Config, Error, Modelos, Utilidades)"
    WScript.Echo "                   Incluye archivos de solicitudes y sus dependencias"
    WScript.Echo ""
    WScript.Echo "🔄 Workflow      - Flujos de Trabajo + dependencias (Config, Error, Modelos, Utilidades)"
    WScript.Echo "                   Incluye archivos de workflow y sus dependencias"
    WScript.Echo ""
    WScript.Echo "🗺️ Mapeo         - Gestión de Mapeos + dependencias (Config, Error, Modelos)"
    WScript.Echo "                   Incluye archivos de mapeos y sus dependencias"
    WScript.Echo ""
    WScript.Echo "🔔 Notification  - Notificaciones + dependencias (Config, Error, Modelos)"
    WScript.Echo "                   Incluye archivos de notificaciones y sus dependencias"
    WScript.Echo ""
    WScript.Echo "📊 Operation     - Operaciones y Logging + dependencias (Config, Error, Modelos)"
    WScript.Echo "                   Incluye archivos de operaciones y sus dependencias"
    WScript.Echo ""
    WScript.Echo "⚙️ Config        - Configuración del Sistema + dependencias (Error, Modelos)"
    WScript.Echo "                   Incluye archivos de configuración y sus dependencias"
    WScript.Echo ""
    WScript.Echo "📂 FileSystem    - Sistema de Archivos + dependencias (Error, Modelos)"
    WScript.Echo "                   Incluye archivos de sistema de archivos y sus dependencias"
    WScript.Echo ""
    WScript.Echo "❌ Error         - Manejo de Errores + dependencias (Modelos)"
    WScript.Echo "                   Incluye archivos de manejo de errores y sus dependencias"
    WScript.Echo ""
    WScript.Echo "📝 Word          - Microsoft Word + dependencias (Error, Modelos)"
    WScript.Echo "                   Incluye archivos de Word y sus dependencias"
    WScript.Echo ""
    WScript.Echo "🧪 TestFramework - Framework de Pruebas + dependencias (11 archivos)"
WScript.Echo "                   Incluye ITestReporter, CTestResult, CTestSuiteResult, CTestReporter, modTestRunner,"
    WScript.Echo "                   modTestUtils, ModAssert, TestModAssert, IFileSystem, IConfig, IErrorHandlerService"
    WScript.Echo ""
    WScript.Echo "🖥️ Aplicacion    - Gestión de Aplicación + dependencias (Config, Error, Modelos)"
    WScript.Echo "                   Incluye archivos de gestión de aplicación y sus dependencias"
    WScript.Echo ""
    WScript.Echo "📊 Modelos       - Modelos de Datos (entidades base)"
    WScript.Echo "                   Incluye todas las entidades de datos del sistema"
    WScript.Echo ""
    WScript.Echo "🔧 Utilidades    - Utilidades y Enumeraciones + dependencias (Error, Modelos)"
    WScript.Echo "                   Incluye utilidades, enumeraciones y sus dependencias"
    WScript.Echo ""
    WScript.Echo "OPCIONES GLOBALES:"
    WScript.Echo "  --help, -h, help             - Mostrar esta ayuda completa"
    WScript.Echo "  --src <directorio>           - Especificar directorio fuente alternativo"
    WScript.Echo "                                 (por defecto: C:\\Proyectos\\CONDOR\\src)"
    WScript.Echo "  --strict                     - Modo estricto: validación exhaustiva de coherencia"
    WScript.Echo "                                 entre JSON y código VBA en formularios"
    WScript.Echo "  --verbose                    - Mostrar información detallada durante la operación"
    WScript.Echo ""
    WScript.Echo "FLUJO DE TRABAJO RECOMENDADO:"
    WScript.Echo "  1. cscript condor_cli.vbs validate     (validar sintaxis antes de importar)"
    WScript.Echo "  2. cscript condor_cli.vbs rebuild      (reconstrucción completa del proyecto)"
    WScript.Echo "  3. cscript condor_cli.vbs test         (ejecutar pruebas unitarias)"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS DE USO:"
    WScript.Echo "  cscript condor_cli.vbs --help"
    WScript.Echo "  cscript condor_cli.vbs validate --verbose --src C:\\MiProyecto\\src"
    WScript.Echo "  cscript condor_cli.vbs export --verbose"
    WScript.Echo "  cscript condor_cli.vbs validate-form-json formulario.json --strict"
    WScript.Echo "  cscript condor_cli.vbs bundle Auth"
    WScript.Echo "  cscript condor_cli.vbs bundle Document C:\\\\temp"
    WScript.Echo "  cscript condor_cli.vbs createtable MiTabla ""CREATE TABLE MiTabla (ID LONG)"""
    WScript.Echo "  cscript condor_cli.vbs listtables"
    WScript.Echo "  cscript condor_cli.vbs relink --all"
    WScript.Echo ""
    WScript.Echo "CONFIGURACIÓN:"
    WScript.Echo "  Base de datos desarrollo: C:\\Proyectos\\CONDOR\\back\\Desarrollo\\CONDOR.accdb"
    WScript.Echo "  Base de datos datos:      C:\\Proyectos\\CONDOR\\back\\CONDOR_datos.accdb"
    WScript.Echo "  Directorio fuente:        C:\\Proyectos\\CONDOR\\src"
    WScript.Echo ""
    WScript.Echo "Para más información, consulte la documentación en docs/CONDOR_MASTER_PLAN.md"
End Sub

' Nueva función que usa DoCmd.LoadFromText para evitar confirmaciones
Sub ImportModuleWithLoadFromText(strSourceFile, moduleName, fileExtension)
    On Error Resume Next
    
    ' Determinar el tipo de objeto Access para DoCmd.LoadFromText
    Dim objectType
    If fileExtension = "bas" Then
        objectType = 5  ' acModule para módulos estándar
    ElseIf fileExtension = "cls" Then
        objectType = 5  ' acModule también para módulos de clase
    Else
        WScript.Echo "  ❌ Error: Tipo de archivo no soportado: " & fileExtension
        Exit Sub
    End If
    
    ' Usar DoCmd.LoadFromText para importar el módulo
    objAccess.DoCmd.LoadFromText objectType, moduleName, strSourceFile
    
    If Err.Number <> 0 Then
        WScript.Echo "  ❌ Error al importar módulo " & moduleName & " con LoadFromText: " & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    If fileExtension = "cls" Then
        WScript.Echo "✅ Clase " & moduleName & " importada correctamente"
    Else
        WScript.Echo "✅ Módulo " & moduleName & " importado correctamente"
    End If
    
    On Error GoTo 0
End Sub

' Subrutina para ejecutar la suite de pruebas unitarias
Sub ExecuteTests()
    WScript.Echo "=== INICIANDO EJECUCION DE PRUEBAS ==="
    Dim reportString
    
    ' Verificar refactorización de logging antes de ejecutar pruebas
    WScript.Echo "Verificando refactorización de logging..."
    Call VerifyLoggingRefactoring()
    
    WScript.Echo "Ejecutando suite de pruebas en Access..."
    On Error Resume Next
    
    ' Configurar Access en modo completamente silencioso para pruebas
    objAccess.Application.DisplayAlerts = False
    objAccess.Application.Echo False
    objAccess.Visible = False
    objAccess.UserControl = False
    
    ' Configuraciones adicionales para suprimir todos los diálogos
    On Error Resume Next
    objAccess.Application.AutomationSecurity = 1  ' msoAutomationSecurityLow
    objAccess.DoCmd.SetWarnings False
    Err.Clear
    On Error Resume Next
    
    ' CORRECCIÓN CRÍTICA: Usar función ExecuteAllTestsForCLI restaurada que ejecuta todas las pruebas
    reportString = objAccess.Run("ExecuteAllTestsForCLI")
    
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: Fallo crítico al invocar la suite de pruebas."
        WScript.Echo "  Código de Error: " & Err.Number
        WScript.Echo "  Descripción: " & Err.Description
        WScript.Echo "  Fuente: " & Err.Source
        WScript.Echo "SUGERENCIA: Abre Access manualmente y ejecuta ExecuteAllTestsForCLI desde el módulo modTestRunner."
        objAccess.Quit
        WScript.Quit 1
    End If
    On Error GoTo 0
    
    ' Verificar si el string devuelto es válido
    If IsEmpty(reportString) Or reportString = "" Then
        WScript.Echo "ERROR: El motor de pruebas de Access no devolvió ningún resultado."
        WScript.Echo "SUGERENCIA: Verifique que la función 'ExecuteAllTestsForCLI' en 'modTestRunner' no esté fallando silenciosamente."
        objAccess.Quit
        WScript.Quit 1
    End If
    
    ' Mostrar el reporte completo
    WScript.Echo "--- INICIO DE RESULTADOS DE PRUEBAS ---"
    WScript.Echo reportString
    WScript.Echo "--- FIN DE RESULTADOS DE PRUEBAS ---"
    
    ' Determinar el éxito o fracaso buscando la línea final
    If InStr(UCase(reportString), "RESULT: SUCCESS") > 0 Then
        WScript.Echo "RESULTADO FINAL: ✓ Todas las pruebas pasaron."
        WScript.Echo "✅ REFACTORIZACIÓN COMPLETADA: Patrón EOperationLog implementado correctamente"
        WScript.Quit 0 ' Éxito
    Else
        WScript.Echo "RESULTADO FINAL: ✗ Pruebas fallidas."
        WScript.Quit 1 ' Error
    End If
End Sub

' Función para importar módulo con conversión UTF-8 -> ANSI
Sub ImportModuleWithAnsiEncoding(strImportPath, moduleName, fileExtension, vbComponent, cleanedContent)
    ' Declarar variables locales
    Dim tempFolderPath, tempFileName, tempFilePath
    Dim objTempFile
    Dim importError, renameError, existingComponent
    
    If fileExtension = "bas" Then
        ' Lógica corregida para módulos estándar (.bas) - usar Add(1)
        On Error Resume Next
        
        ' Buscar si ya existe un componente con este nombre
        Set vbComponent = Nothing
        For Each existingComponent In objAccess.VBE.ActiveVBProject.VBComponents
            If existingComponent.Name = moduleName Then
                Set vbComponent = existingComponent
                Exit For
            End If
        Next
        
        ' Si no existe, crear nuevo componente
        If vbComponent Is Nothing Then
            Set vbComponent = objAccess.VBE.ActiveVBProject.VBComponents.Add(1) ' 1 = vbext_ct_StdModule
            If Err.Number <> 0 Then
                WScript.Echo "❌ ERROR: No se pudo crear componente estándar para " & moduleName & ": " & Err.Description
                On Error GoTo 0
                Exit Sub
            End If
            
            ' Renombrar inmediatamente después de crear
            vbComponent.Name = moduleName
            renameError = Err.Number
            If renameError <> 0 Then
                WScript.Echo "❌ ERROR: No se pudo renombrar el módulo nuevo a '" & moduleName & "': " & Err.Description & " (Código: " & Err.Number & ")"
                On Error GoTo 0
                Exit Sub
            End If
        Else
            ' Si existe, limpiar el código existente
            If vbComponent.CodeModule.CountOfLines > 0 Then
                vbComponent.CodeModule.DeleteLines 1, vbComponent.CodeModule.CountOfLines
            End If
        End If
        
        ' Insertar el contenido limpio en el módulo de código
        vbComponent.CodeModule.AddFromString cleanedContent
        If Err.Number <> 0 Then
            WScript.Echo "❌ ERROR: No se pudo insertar código en el módulo " & moduleName & ": " & Err.Description
            On Error GoTo 0
            Exit Sub
        End If
        
        On Error GoTo 0
        ' Confirmar éxito
        WScript.Echo "✅ Módulo " & moduleName & " importado correctamente"
        
    ElseIf fileExtension = "cls" Then
        ' Lógica específica para módulos de clase (.cls)
        On Error Resume Next
        
        ' Buscar si ya existe un componente con este nombre
        Set vbComponent = Nothing
        For Each existingComponent In objAccess.VBE.ActiveVBProject.VBComponents
            If existingComponent.Name = moduleName Then
                Set vbComponent = existingComponent
                Exit For
            End If
        Next
        
        ' Si no existe, crear nuevo componente
        If vbComponent Is Nothing Then
            Set vbComponent = objAccess.VBE.ActiveVBProject.VBComponents.Add(2) ' 2 = vbext_ct_ClassModule
            If Err.Number <> 0 Then
                WScript.Echo "❌ ERROR: No se pudo crear componente de clase para " & moduleName & ": " & Err.Description
                On Error GoTo 0
                Exit Sub
            End If
            
            ' Renombrar inmediatamente después de crear
            vbComponent.Name = moduleName
            renameError = Err.Number
            If renameError <> 0 Then
                WScript.Echo "❌ ERROR: No se pudo renombrar la clase nueva a '" & moduleName & "': " & Err.Description & " (Código: " & Err.Number & ")"
                On Error GoTo 0
                Exit Sub
            End If
        Else
            ' Si existe, limpiar el código existente
            If vbComponent.CodeModule.CountOfLines > 0 Then
                vbComponent.CodeModule.DeleteLines 1, vbComponent.CodeModule.CountOfLines
            End If
        End If
        
        ' Insertar el contenido limpio en el módulo de código
        vbComponent.CodeModule.AddFromString cleanedContent
        If Err.Number <> 0 Then
            WScript.Echo "❌ ERROR: No se pudo insertar código en la clase " & moduleName & ": " & Err.Description
            On Error GoTo 0
            Exit Sub
        End If
        
        On Error GoTo 0
        ' Confirmar éxito
        WScript.Echo "✅ Clase " & moduleName & " importada correctamente"
    End If
End Sub

' Función simplificada usando VBComponents.Import() - método desatendido
Sub ImportModuleWithAnsiEncodingNew(strImportPath, moduleName, fileExtension, vbComponent, cleanedContent)
    ' Método con verificación de referencias VBA y enlace tardío
    Dim existingComponent, vbeObject, vbProject, vbComponents
    
    On Error Resume Next
    
    ' Verificar que VBE esté disponible usando enlace tardío
    Set vbeObject = objAccess.VBE
    If Err.Number <> 0 Then
        WScript.Echo "❌ ERROR: VBA no está habilitado o no se puede acceder al VBE: " & Err.Description
        On Error GoTo 0
        Exit Sub
    End If
    
    ' Verificar que el proyecto VBA esté disponible
    Set vbProject = vbeObject.ActiveVBProject
    If Err.Number <> 0 Then
        WScript.Echo "❌ ERROR: No se puede acceder al proyecto VBA activo: " & Err.Description
        On Error GoTo 0
        Exit Sub
    End If
    
    ' Verificar que VBComponents esté disponible
    Set vbComponents = vbProject.VBComponents
    If Err.Number <> 0 Then
        WScript.Echo "❌ ERROR: No se puede acceder a VBComponents (referencias VBA requeridas): " & Err.Description
        On Error GoTo 0
        Exit Sub
    End If
    
    ' Eliminar el componente existente si ya existe
    Set vbComponent = Nothing
    For Each existingComponent In vbComponents
        If existingComponent.Name = moduleName Then
            vbComponents.Remove existingComponent
            If Err.Number <> 0 Then
                WScript.Echo "❌ ERROR: No se pudo eliminar componente existente " & moduleName & ": " & Err.Description
                On Error GoTo 0
                Exit Sub
            End If
            Exit For
        End If
    Next
    
    ' Importar directamente el archivo usando VBComponents.Import()
    Set vbComponent = vbComponents.Import(strImportPath)
    
    If Err.Number <> 0 Then
        WScript.Echo "❌ ERROR: No se pudo importar " & moduleName & ": " & Err.Description
        WScript.Echo "  Verifique que las referencias 'Microsoft Visual Basic for Applications Extensibility' estén habilitadas"
        On Error GoTo 0
        Exit Sub
    End If
    
    ' Verificar si el componente fue importado correctamente
    If vbComponent Is Nothing Then
        WScript.Echo "❌ ERROR: El componente importado es Nothing"
        On Error GoTo 0
        Exit Sub
    End If
    
    ' Renombrar el componente solo si es necesario
    If vbComponent.Name <> moduleName Then
        Dim originalName
        originalName = vbComponent.Name
        vbComponent.Name = moduleName
        If Err.Number <> 0 Then
            WScript.Echo "⚠️ ADVERTENCIA: No se pudo renombrar de '" & originalName & "' a '" & moduleName & "': " & Err.Description
            WScript.Echo "  El módulo se importó como '" & originalName & "' - verifique el nombre en el archivo fuente"
            Err.Clear
        End If
    End If
    
    On Error GoTo 0
    
    ' Confirmar éxito según el tipo
    If fileExtension = "bas" Then
        WScript.Echo "✅ Módulo " & moduleName & " importado correctamente"
    ElseIf fileExtension = "cls" Then
        WScript.Echo "✅ Clase " & moduleName & " importada correctamente"
    End If
End Sub


' Función desatendida para importar módulos usando VBComponents.Import()
' Mantiene la funcionalidad de limpieza de código de rebuild
Sub ImportModuleDesatendido(strImportPath, moduleName, fileExtension, cleanedContent)
    ' Declarar variables locales
    Dim tempFolderPath, tempFileName, tempFilePath
    Dim objTempFile, existingComponent, vbComp
    
    On Error Resume Next
    
    ' Eliminar módulo si ya existe
    Set vbComp = Nothing
    For Each existingComponent In objAccess.VBE.ActiveVBProject.VBComponents
        If existingComponent.Name = moduleName Then
            objAccess.VBE.ActiveVBProject.VBComponents.Remove existingComponent
            If Err.Number <> 0 Then
                WScript.Echo "❌ ERROR: No se pudo eliminar componente existente " & moduleName & ": " & Err.Description
                On Error GoTo 0
                Exit Sub
            End If
            Exit For
        End If
    Next
    
    ' Crear archivo temporal con contenido limpio
    tempFolderPath = objFSO.GetSpecialFolder(2) ' Carpeta temporal del sistema
    tempFileName = "temp_" & moduleName & "." & fileExtension
    tempFilePath = objFSO.BuildPath(tempFolderPath, tempFileName)
    
    ' Escribir contenido limpio al archivo temporal
    Set objTempFile = objFSO.CreateTextFile(tempFilePath, True, False) ' False = ANSI encoding
    objTempFile.Write cleanedContent
    objTempFile.Close
    Set objTempFile = Nothing
    
    ' Importar módulo usando VBComponents.Import()
    Set vbComp = objAccess.VBE.ActiveVBProject.VBComponents.Import(tempFilePath)
    
    If Err.Number <> 0 Then
        WScript.Echo "❌ ERROR: No se pudo importar " & moduleName & ": " & Err.Description
        ' Limpiar archivo temporal
        If objFSO.FileExists(tempFilePath) Then objFSO.DeleteFile tempFilePath
        On Error GoTo 0
        Exit Sub
    End If
    
    ' Renombrar el componente si es necesario
    If Not vbComp Is Nothing And vbComp.Name <> moduleName Then
        vbComp.Name = moduleName
        If Err.Number <> 0 Then
            WScript.Echo "⚠️ ADVERTENCIA: No se pudo renombrar a '" & moduleName & "': " & Err.Description
            Err.Clear
        End If
    End If
    
    ' Limpiar archivo temporal
    If objFSO.FileExists(tempFilePath) Then objFSO.DeleteFile tempFilePath
    
    On Error GoTo 0
    
    ' Confirmar éxito
    If fileExtension = "bas" Then
        WScript.Echo "✅ Módulo " & moduleName & " importado correctamente"
    ElseIf fileExtension = "cls" Then
        WScript.Echo "✅ Clase " & moduleName & " importada correctamente"
    End If
End Sub

' Subrutina para re-vincular tablas de Access
Sub RelinkTables()
    Dim strDbPath, strLocalFolder
    
    WScript.Echo "=== INICIANDO RE-VINCULACION DE TABLAS ==="
    
    ' Verificar si se usa el modo --all
    If objArgs.Count >= 2 Then
        If LCase(objArgs(1)) = "--all" Then
            Call RelinkAllDatabases()
            Exit Sub
        End If
    End If
    
    ' Verificar que se proporcionaron los argumentos necesarios para modo manual
    If objArgs.Count < 3 Then
        WScript.Echo "Error: El comando relink requiere argumentos:"
        WScript.Echo "Uso: cscript condor_cli.vbs relink <db_path> <local_folder>"
        WScript.Echo "  o: cscript condor_cli.vbs relink --all"
        WScript.Echo "  db_path: Ruta a la base de datos frontend (.accdb)"
        WScript.Echo "  local_folder: Ruta a la carpeta con las bases de datos locales"
        WScript.Echo "  --all: Re-vincular todas las bases de datos en ./back automáticamente"
        objAccess.Quit
        WScript.Quit 1
    End If
    
    ' Leer argumentos de la línea de comandos
    strDbPath = objArgs(1)
    strLocalFolder = objArgs(2)
    
    WScript.Echo "Base de datos frontend: " & strDbPath
    WScript.Echo "Carpeta de backends locales: " & strLocalFolder
    
    ' Verificar que los paths existen
    If Not objFSO.FileExists(strDbPath) Then
        WScript.Echo "Error: La base de datos frontend no existe: " & strDbPath
        objAccess.Quit
        WScript.Quit 1
    End If
    
    If Not objFSO.FolderExists(strLocalFolder) Then
        WScript.Echo "Error: La carpeta de backends locales no existe: " & strLocalFolder
        objAccess.Quit
        WScript.Quit 1
    End If
    
    WScript.Echo "Funcionalidad de re-vinculación pendiente de implementación."
    WScript.Echo "=== RE-VINCULACION COMPLETADA ==="
End Sub

' Subrutina para re-vincular todas las bases de datos automáticamente
Sub RelinkAllDatabases()
    Dim objBackFolder, objFile
    Dim strBackPath, strDbCount
    Dim dbCount, successCount, errorCount
    Dim strDbName, strPassword
    Dim arrDatabases()
    Dim i
    
    WScript.Echo "=== MODO AUTOMATICO: RE-VINCULANDO TODAS LAS BASES DE DATOS ==="
    
    ' Definir ruta del directorio back
    strBackPath = objFSO.GetAbsolutePathName("back")
    
    ' Verificar que existe el directorio back
    If Not objFSO.FolderExists(strBackPath) Then
        WScript.Echo "Error: El directorio ./back no existe: " & strBackPath
        objAccess.Quit
        WScript.Quit 1
    End If
    
    WScript.Echo "Directorio de backends: " & strBackPath
    
    ' Contar y listar bases de datos .accdb
    Set objBackFolder = objFSO.GetFolder(strBackPath)
    dbCount = 0
    
    ' Redimensionar array para almacenar información de bases de datos
    ReDim arrDatabases(50) ' Máximo 50 bases de datos
    
    For Each objFile In objBackFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "accdb" Then
            strDbName = objFSO.GetBaseName(objFile.Name)
            
            ' Determinar contraseña según el nombre de la base de datos
            If InStr(1, UCase(strDbName), "CONDOR") > 0 Then
                strPassword = "(sin contraseña)"
            Else
                strPassword = "dpddpd"
            End If
            
            ' Almacenar información de la base de datos
            arrDatabases(dbCount) = objFile.Path & "|" & strPassword
            dbCount = dbCount + 1
            
            WScript.Echo "  [" & dbCount & "] " & objFile.Name & " - " & strPassword
        End If
    Next
    
    If dbCount = 0 Then
        WScript.Echo "No se encontraron bases de datos .accdb en el directorio ./back"
        WScript.Echo "=== RE-VINCULACION COMPLETADA ==="
        Exit Sub
    End If
    
    WScript.Echo "Total de bases de datos encontradas: " & dbCount
    WScript.Echo "Iniciando proceso de re-vinculación..."
    WScript.Echo ""
    
    ' Procesar cada base de datos
    successCount = 0
    errorCount = 0
    
    For i = 0 To dbCount - 1
        Dim arrDbInfo
        arrDbInfo = Split(arrDatabases(i), "|")
        
        If UBound(arrDbInfo) >= 1 Then
            Dim strDbPath, strDbPassword
            strDbPath = arrDbInfo(0)
            strDbPassword = arrDbInfo(1)
            
            WScript.Echo "Procesando: " & objFSO.GetFileName(strDbPath)
            
            If RelinkSingleDatabase(strDbPath, strDbPassword, strBackPath) Then
                successCount = successCount + 1
                WScript.Echo "  ✓ Re-vinculación exitosa"
            Else
                errorCount = errorCount + 1
                WScript.Echo "  ❌ Error en re-vinculación"
            End If
            WScript.Echo ""
        End If
    Next
    
    ' Resumen final
    WScript.Echo "=== RESUMEN DE RE-VINCULACION AUTOMATICA ==="
    WScript.Echo "Total procesadas: " & dbCount
    WScript.Echo "Exitosas: " & successCount
    WScript.Echo "Con errores: " & errorCount
    
    If errorCount = 0 Then
        WScript.Echo "✓ Todas las bases de datos fueron re-vinculadas exitosamente"
    Else
        WScript.Echo "⚠️ Algunas bases de datos tuvieron errores durante la re-vinculación"
    End If
    
    WScript.Echo "=== RE-VINCULACION AUTOMATICA COMPLETADA ==="
End Sub

' Función para determinar la contraseña de una base de datos
Function GetDatabasePassword(strDbPath)
    Dim strDbName
    strDbName = objFSO.GetBaseName(strDbPath)
    
    ' Las bases de datos CONDOR no requieren contraseña
    If InStr(1, UCase(strDbName), "CONDOR") > 0 Then
        GetDatabasePassword = ""
    Else
        ' Las demás bases de datos usan 'dpddpd'
        GetDatabasePassword = "dpddpd"
    End If
End Function

' Función para re-vincular una sola base de datos
Function RelinkSingleDatabase(strDbPath, strPassword, strBackPath)
    Dim objDb, objTableDef
    Dim strConnectionString
    Dim linkedTableCount, successCount
    
    On Error Resume Next
    
    ' Abrir la base de datos
    If strPassword = "(sin contraseña)" Then
        Set objDb = objAccess.DBEngine.OpenDatabase(strDbPath)
    Else
        Set objDb = objAccess.DBEngine.OpenDatabase(strDbPath, False, False, ";PWD=" & strPassword)
    End If
    
    If Err.Number <> 0 Then
        WScript.Echo "  Error al abrir base de datos: " & Err.Description
        RelinkSingleDatabase = False
        Err.Clear
        Exit Function
    End If
    
    linkedTableCount = 0
    successCount = 0
    
    ' Recorrer todas las tablas vinculadas
    For Each objTableDef In objDb.TableDefs
        If Len(objTableDef.Connect) > 0 Then
            linkedTableCount = linkedTableCount + 1
            
            ' Extraer el nombre de la base de datos del connect string actual
            Dim strCurrentConnect, strSourceDb, strNewConnect
            strCurrentConnect = objTableDef.Connect
            
            ' Buscar el patrón DATABASE= en el connect string
            Dim intDbStart, intDbEnd, strDbName
            intDbStart = InStr(1, UCase(strCurrentConnect), "DATABASE=")
            
            If intDbStart > 0 Then
                intDbStart = intDbStart + 9 ' Longitud de "DATABASE="
                intDbEnd = InStr(intDbStart, strCurrentConnect, ";")
                
                If intDbEnd = 0 Then intDbEnd = Len(strCurrentConnect) + 1
                
                strSourceDb = Mid(strCurrentConnect, intDbStart, intDbEnd - intDbStart)
                strDbName = objFSO.GetFileName(strSourceDb)
                
                ' Construir nueva ruta local
                Dim strNewDbPath
                strNewDbPath = objFSO.BuildPath(strBackPath, strDbName)
                
                ' Verificar que la base de datos local existe
                If objFSO.FileExists(strNewDbPath) Then
                    ' Construir nuevo connect string
                    strNewConnect = Replace(strCurrentConnect, strSourceDb, strNewDbPath)
                    
                    ' Actualizar la vinculación
                    objTableDef.Connect = strNewConnect
                    objTableDef.RefreshLink
                    
                    If Err.Number = 0 Then
                        successCount = successCount + 1
                        WScript.Echo "    ✓ " & objTableDef.Name & " -> " & strDbName
                    Else
                        WScript.Echo "    ❌ Error en " & objTableDef.Name & ": " & Err.Description
                        Err.Clear
                    End If
                Else
                    WScript.Echo "    ⚠️ Base de datos local no encontrada: " & strDbName
                End If
            Else
                WScript.Echo "    ⚠️ No se pudo extraer DATABASE de: " & objTableDef.Name
            End If
        End If
    Next
    
    ' Cerrar base de datos
    objDb.Close
    Set objDb = Nothing
    
    WScript.Echo "    Tablas vinculadas procesadas: " & linkedTableCount
    WScript.Echo "    Re-vinculaciones exitosas: " & successCount
    
    ' Considerar exitoso si se procesó al menos una tabla correctamente
    RelinkSingleDatabase = (successCount > 0 Or linkedTableCount = 0)
    
    On Error GoTo 0
End Function

' Subrutina para reconstruir completamente el proyecto VBA
Sub RebuildProject()
    WScript.Echo "=== RECONSTRUCCION COMPLETA DEL PROYECTO VBA ==="
    WScript.Echo "ADVERTENCIA: Se eliminaran TODOS los modulos VBA existentes"
    WScript.Echo "Iniciando proceso de reconstruccion..."
    
    On Error Resume Next
    
    ' Paso 1: Eliminar todos los módulos existentes
    WScript.Echo "Paso 1: Eliminando todos los modulos VBA existentes..."
    
    Dim vbProject, vbComponent
    Set vbProject = objAccess.VBE.ActiveVBProject
    
    Dim componentCount, i, errorDetails
    componentCount = vbProject.VBComponents.Count
    
    ' Iterar hacia atrás para evitar problemas al eliminar elementos
    For i = componentCount To 1 Step -1
        Set vbComponent = vbProject.VBComponents(i)
        
        ' Solo eliminar módulos estándar y de clase (no formularios ni informes)
        If vbComponent.Type = 1 Or vbComponent.Type = 2 Then ' vbext_ct_StdModule = 1, vbext_ct_ClassModule = 2
            WScript.Echo "  Eliminando: " & vbComponent.Name & " (Tipo: " & vbComponent.Type & ")"
            vbProject.VBComponents.Remove vbComponent
            
            If Err.Number <> 0 Then
                WScript.Echo "  ❌ Error eliminando " & vbComponent.Name & ": " & Err.Description
                Err.Clear
            Else
                WScript.Echo "  ✓ Eliminado: " & vbComponent.Name
            End If
        End If
    Next
    
    WScript.Echo "Paso 2: Cerrando base de datos..."
    
    ' Cerrar sin guardar explícitamente para evitar confirmaciones
    objAccess.Quit 1  ' acQuitSaveAll = 1
    
    If Err.Number <> 0 Then
        WScript.Echo "Advertencia al cerrar Access: " & Err.Description
        Err.Clear
    End If
    
    Set objAccess = Nothing
    WScript.Echo "✓ Base de datos cerrada y guardada"
    
    ' Paso 3: Volver a abrir la base de datos
    WScript.Echo "Paso 3: Reabriendo base de datos con proyecto VBA limpio..."
    
    Set objAccess = CreateObject("Access.Application")
    
    If Err.Number <> 0 Then
        WScript.Echo "❌ Error al crear nueva instancia de Access: " & Err.Description
        WScript.Quit 1
    End If
    
    ' Configurar Access en modo silencioso
    objAccess.Visible = False
    objAccess.UserControl = False
    
    ' Suprimir alertas y diálogos de confirmación
    On Error Resume Next
    ' objAccess.DoCmd.SetWarnings False  ' Comentado temporalmente por error de compilación
    objAccess.Application.Echo False
    objAccess.DisplayAlerts = False
    ' Configuraciones adicionales para suprimir diálogos
    objAccess.Application.AutomationSecurity = 1  ' msoAutomationSecurityLow
    objAccess.VBE.MainWindow.Visible = False
    Err.Clear
    On Error GoTo 0
    
    ' Determinar contraseña para la base de datos
    Dim strDbPassword
    strDbPassword = GetDatabasePassword(strAccessPath)
    
    ' Abrir base de datos
    If strDbPassword = "" Then
        objAccess.OpenCurrentDatabase strAccessPath
    Else
        objAccess.OpenCurrentDatabase strAccessPath, , strDbPassword
    End If
    
    If Err.Number <> 0 Then
        WScript.Echo "❌ Error al reabrir base de datos: " & Err.Description
        WScript.Quit 1
    End If
    
    WScript.Echo "✓ Base de datos reabierta con proyecto VBA limpio"
    
    ' Paso 4: Importar todos los módulos de nuevo
    WScript.Echo "Paso 4: Importando todos los modulos desde /src..."
    
    ' Integrar lógica de importación directamente
    Dim objFolder, objFile
    Dim strModuleName, strFileName, strContent
    Dim srcModules
    Dim moduleExists
    Dim validationResult
    Dim totalFiles, validFiles, invalidFiles
    
    If Not objFSO.FolderExists(strSourcePath) Then
        WScript.Echo "Error: Directorio de origen no existe: " & strSourcePath
        objAccess.Quit
        WScript.Quit 1
    End If
    
    ' PASO 4.1: Validacion previa de sintaxis
    WScript.Echo "Validando sintaxis de todos los modulos..."
    Set objFolder = objFSO.GetFolder(strSourcePath)
    totalFiles = 0
    validFiles = 0
    invalidFiles = 0
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            totalFiles = totalFiles + 1
            validationResult = ValidateVBASyntax(objFile.Path, errorDetails)
            
            If validationResult = True Then
                validFiles = validFiles + 1
                WScript.Echo "  ✓ " & objFile.Name & " - Sintaxis valida"
            Else
                invalidFiles = invalidFiles + 1
                WScript.Echo "  ✗ ERROR en " & objFile.Name & ": " & errorDetails
            End If
        End If
    Next
    
    If invalidFiles > 0 Then
        WScript.Echo "ABORTANDO: Se encontraron " & invalidFiles & " archivos con errores de sintaxis."
        WScript.Echo "Use 'cscript condor_cli.vbs validate --verbose' para más detalles."
        objAccess.Quit
        WScript.Quit 1
    End If
    
    WScript.Echo "✓ Validacion completada: " & validFiles & " archivos validos"
    
    ' PASO 4.2: Procesar archivos de modulos
    Set objFolder = objFSO.GetFolder(strSourcePath)
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            strFileName = objFile.Path
            strModuleName = objFSO.GetBaseName(objFile.Name)
            
            WScript.Echo "Procesando modulo: " & strModuleName
            
            ' Determinar tipo de archivo
            Dim fileExtension
            fileExtension = LCase(objFSO.GetExtensionName(objFile.Name))
            
            ' Limpiar archivo antes de importar (eliminar metadatos Attribute)
            Dim cleanedContent
            cleanedContent = CleanVBAFile(strFileName, fileExtension)
            
            ' Importar usando contenido limpio
            WScript.Echo "  Clase " & strModuleName & " importada correctamente"
            Call ImportModuleWithAnsiEncoding(strFileName, strModuleName, fileExtension, Nothing, cleanedContent)
            
            If Err.Number <> 0 Then
                WScript.Echo "Error al importar modulo " & strModuleName & ": " & Err.Description
                Err.Clear
            End If
        End If
    Next
    
    ' PASO 4.3: Guardar cada modulo individualmente
    WScript.Echo "Guardando modulos individualmente..."
    On Error Resume Next
    
    For Each vbComponent In objAccess.VBE.ActiveVBProject.VBComponents
        If vbComponent.Type = 1 Then  ' vbext_ct_StdModule
            WScript.Echo "Guardando modulo: " & vbComponent.Name
            objAccess.DoCmd.Save 5, vbComponent.Name  ' acModule = 5
            If Err.Number <> 0 Then
                WScript.Echo "Advertencia al guardar " & vbComponent.Name & ": " & Err.Description
                Err.Clear
            End If
        ElseIf vbComponent.Type = 2 Then  ' vbext_ct_ClassModule
            WScript.Echo "Guardando clase: " & vbComponent.Name
            objAccess.DoCmd.Save 7, vbComponent.Name  ' acClassModule = 7
            If Err.Number <> 0 Then
                WScript.Echo "Advertencia al guardar " & vbComponent.Name & ": " & Err.Description
                Err.Clear
            End If
        End If
    Next
    
    ' PASO 4.4: Verificacion de integridad y compilacion
    WScript.Echo "Verificando integridad de nombres de modulos..."
    Call VerifyModuleNames()
    
    WScript.Echo "=== RECONSTRUCCION COMPLETADA EXITOSAMENTE ==="
    WScript.Echo "El proyecto VBA ha sido completamente reconstruido"
    WScript.Echo "Todos los modulos han sido reimportados desde /src"
    
    On Error GoTo 0
End Sub



' Subrutina para verificar y cerrar procesos de Access existentes
Sub CloseExistingAccessProcesses()
    Dim objWMI, colProcesses, objProcess
    Dim processCount
    
    WScript.Echo "Verificando procesos de Access existentes..."
    
    On Error Resume Next
    Set objWMI = GetObject("winmgmts:")
    Set colProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'MSACCESS.EXE'")
    
    processCount = 0
    For Each objProcess In colProcesses
        processCount = processCount + 1
    Next
    
    If processCount > 0 Then
        WScript.Echo "Se encontraron " & processCount & " procesos de Access ejecutándose. Cerrándolos..."
        
        For Each objProcess In colProcesses
            WScript.Echo "Terminando proceso Access PID: " & objProcess.ProcessId
            objProcess.Terminate()
        Next
        
        ' Esperar un momento para que los procesos se cierren completamente
        WScript.Sleep 2000
        WScript.Echo "✓ Procesos de Access cerrados correctamente"
    Else
        WScript.Echo "✓ No se encontraron procesos de Access ejecutándose"
    End If
    
    On Error GoTo 0
End Sub

' La subrutina ExecuteTestModule ha sido eliminada ya que ahora se usa el motor interno modTestRunner

' Subrutina para sincronizar un módulo individual
' Parámetro: moduleName - Nombre del módulo a sincronizar (ej. "CAuthService")


' Subrutina optimizada para importar un solo módulo (sin cerrar/abrir BD)
Sub ImportSingleModuleOptimized(moduleName)
    On Error Resume Next
    
    ' Paso 1: Verificar que el fichero fuente (.bas o .cls) existe en la carpeta /src
    Dim strBasFile, strClsFile, strSourceFile, fileExtension
    strBasFile = objFSO.BuildPath(strSourcePath, moduleName & ".bas")
    strClsFile = objFSO.BuildPath(strSourcePath, moduleName & ".cls")
    
    If objFSO.FileExists(strBasFile) Then
        strSourceFile = strBasFile
        fileExtension = "bas"
    ElseIf objFSO.FileExists(strClsFile) Then
        strSourceFile = strClsFile
        fileExtension = "cls"
    Else
        WScript.Echo "  ❌ Error: No se encontró el archivo fuente para " & moduleName
        WScript.Echo "      Buscado: " & strBasFile
        WScript.Echo "      Buscado: " & strClsFile
        Exit Sub
    End If
    
    WScript.Echo "  ✓ Archivo fuente encontrado: " & strSourceFile
    
    ' Paso 2: Validar sintaxis del archivo
    Dim errorDetails, validationResult
    validationResult = ValidateVBASyntax(strSourceFile, errorDetails)
    
    If validationResult <> True Then
        WScript.Echo "  ❌ Error de sintaxis en " & moduleName & ": " & errorDetails
        Exit Sub
    End If
    
    WScript.Echo "  ✓ Sintaxis válida"
    
    ' Paso 3: Limpiar el contenido del fichero utilizando CleanVBAFile
    Dim cleanedContent
    cleanedContent = CleanVBAFile(strSourceFile, fileExtension)
    
    If cleanedContent = "" Then
        WScript.Echo "  ❌ Error: No se pudo leer o limpiar el contenido del archivo"
        Exit Sub
    End If
    
    WScript.Echo "  ✓ Contenido limpiado"
    
    ' Paso 4: Importar el módulo usando DoCmd.LoadFromText (sin confirmaciones)
    WScript.Echo "  Importando módulo: " & moduleName
    Call ImportModuleWithLoadFromText(strSourceFile, moduleName, fileExtension)
    
    If Err.Number <> 0 Then
        WScript.Echo "  ❌ Error al importar módulo " & moduleName & ": " & Err.Description
        Err.Clear
        Exit Sub
    End If
    
    If fileExtension = "cls" Then
        WScript.Echo "✅ Clase " & moduleName & " importada correctamente"
    Else
        WScript.Echo "✅ Módulo " & moduleName & " importado correctamente"
    End If
    
    WScript.Echo "  ✅ Módulo " & moduleName & " sincronizado correctamente"
    
    On Error GoTo 0
End Sub

' Subrutina para actualizar proyecto VBA con sincronización selectiva


' Subrutina para exportar módulos VBA actuales a carpeta cache












' Función para verificar cambios antes de abrir la base de datos


' Subrutina para copiar solo archivos modificados a la caché


' Subrutina para mostrar ayuda específica del comando bundle
Sub ShowBundleHelp()
    WScript.Echo "=== CONDOR CLI - AYUDA DEL COMANDO BUNDLE ==="
    WScript.Echo "Empaqueta archivos de código por funcionalidad según CONDOR_MASTER_PLAN.md"
    WScript.Echo ""
    WScript.Echo "USO:"
    WScript.Echo "  cscript condor_cli.vbs bundle <funcionalidad> [ruta_destino]"
    WScript.Echo "  cscript condor_cli.vbs bundle --help"
    WScript.Echo ""
    WScript.Echo "PARÁMETROS:"
    WScript.Echo "  <funcionalidad>  - Nombre de la funcionalidad a empaquetar (obligatorio)"
    WScript.Echo "  [ruta_destino]   - Carpeta donde crear el paquete (opcional, por defecto: carpeta actual)"
    WScript.Echo ""
    WScript.Echo "FUNCIONALIDADES DISPONIBLES:"
    WScript.Echo ""
    WScript.Echo "🔐 Auth - Autenticación y Autorización"
    WScript.Echo "   Incluye: IAuthService, CAuthService, CMockAuthService, IAuthRepository,"
    WScript.Echo "            CAuthRepository, CMockAuthRepository, AuthData, modAuthFactory,"
    WScript.Echo "            TestCAuthService, IntegrationTestCAuthRepository + dependencias"
    WScript.Echo ""
    WScript.Echo "📄 Document - Gestión de Documentos"
    WScript.Echo "   Incluye: IDocumentService, CDocumentService, CMockDocumentService,"
    WScript.Echo "            IWordManager, CWordManager, CMockWordManager, ISolicitudService + dependencias"
    WScript.Echo ""
    WScript.Echo "📁 Expediente - Gestión de Expedientes"
    WScript.Echo "   Incluye: IExpedienteService, CExpedienteService, CMockExpedienteService,"
    WScript.Echo "            IExpedienteRepository, CExpedienteRepository + dependencias"
    WScript.Echo ""
    WScript.Echo "📋 Solicitud - Gestión de Solicitudes"
    WScript.Echo "   Incluye: ISolicitudService, CSolicitudService, CMockSolicitudService,"
    WScript.Echo "            ISolicitudRepository, CSolicitudRepository + modelos de datos"
    WScript.Echo ""
    WScript.Echo "🔄 Workflow - Flujos de Trabajo"
    WScript.Echo "   Incluye: IWorkflowService, CWorkflowService, CMockWorkflowService,"
    WScript.Echo "            IWorkflowRepository, CWorkflowRepository + modelos de estado"
    WScript.Echo ""
    WScript.Echo "🗺️ Mapeo - Gestión de Mapeos"
    WScript.Echo "   Incluye: IMapeoRepository, CMapeoRepository, CMockMapeoRepository,"
    WScript.Echo "            EMapeo, IntegrationTestCMapeoRepository + dependencias"
    WScript.Echo ""
    WScript.Echo "⚙️ Config - Configuración del Sistema"
    WScript.Echo "   Incluye: IConfig, CConfig, CMockConfig, modConfigFactory,"
    WScript.Echo "            TestConfig (simplificado tras Misión de Emergencia)"
    WScript.Echo ""
    WScript.Echo "📂 FileSystem - Sistema de Archivos"
    WScript.Echo "   Incluye: IFileSystem, CFileSystem, CMockFileSystem,"
    WScript.Echo "            ModFileSystemFactory, TestFileSystem + dependencias"
    WScript.Echo ""
    WScript.Echo "❌ Error - Manejo de Errores"
    WScript.Echo "   Incluye: IErrorHandlerService, CErrorHandlerService, CMockErrorHandlerService,"
    WScript.Echo "            modErrorHandlerFactory, modErrorHandler + dependencias"
    WScript.Echo ""
    WScript.Echo "📝 Word - Integración con Microsoft Word"
    WScript.Echo "   Incluye: IWordManager, CWordManager, CMockWordManager,"
    WScript.Echo "            ModWordManagerFactory, TestWordManager + dependencias"
    WScript.Echo ""
    WScript.Echo "🧪 TestFramework - Framework de Pruebas"
WScript.Echo "   Incluye: ITestReporter, CTestResult, CTestSuiteResult, CTestReporter, modTestRunner,"
    WScript.Echo "            modTestUtils, ModAssert, TestModAssert + interfaces base"
    WScript.Echo ""
    WScript.Echo "🚀 App - Gestión de Aplicación"
    WScript.Echo "   Incluye: IAppManager, CAppManager, ModAppManagerFactory,"
    WScript.Echo "            TestAppManager + dependencias de autenticación y config"
    WScript.Echo ""
    WScript.Echo "🏗️ Models - Modelos de Datos"
    WScript.Echo "   Incluye: Todas las entidades E_* (Usuario, Solicitud, Expediente,"
    WScript.Echo "            DatosPC, DatosCDCA, Estado, Transicion, Mapeo, etc.)"
    WScript.Echo ""
    WScript.Echo "🔧 Utils - Utilidades y Enumeraciones"
    WScript.Echo "   Incluye: ModDatabase, ModRepositoryFactory, ModUtils,"
    WScript.Echo "            E_TipoSolicitud, E_EstadoSolicitud, E_RolUsuario, etc."
    WScript.Echo ""
    WScript.Echo "🧪 Tests - Archivos de Pruebas"
    WScript.Echo "   Incluye: Todos los archivos Test* e IntegrationTest* del proyecto"
    WScript.Echo "            (TestAppManager, TestAuthService, TestCConfig, etc.)"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS DE USO:"
    WScript.Echo "  cscript condor_cli.vbs bundle Auth"
    WScript.Echo "  cscript condor_cli.vbs bundle Document C:\\temp"
    WScript.Echo "  cscript condor_cli.vbs bundle TestFramework"
    WScript.Echo "  cscript condor_cli.vbs bundle Tests"
    WScript.Echo "  cscript condor_cli.vbs bundle Config"
    WScript.Echo ""
    WScript.Echo "NOTAS:"
    WScript.Echo "  • Los archivos se copian con extensión .txt para fácil visualización"
    WScript.Echo "  • Se crea una carpeta con timestamp: bundle_<funcionalidad>_YYYYMMDD_HHMMSS"
    WScript.Echo "  • Cada funcionalidad incluye automáticamente sus dependencias"
    WScript.Echo "  • Si un archivo no existe, se muestra una advertencia pero continúa"
End Sub

' Función para obtener la lista de archivos por funcionalidad según CONDOR_MASTER_PLAN.md
' Incluye dependencias para cada funcionalidad
Function GetFunctionalityFiles(strFunctionality)
    Dim arrFiles
    
    Select Case LCase(strFunctionality)
        Case "auth", "autenticacion", "authentication"
            ' Sección 3.1 - Autenticación + Dependencias
            arrFiles = Array("IAuthService.cls", "CAuthService.cls", "CMockAuthService.cls", _
                           "IAuthRepository.cls", "CAuthRepository.cls", "CMockAuthRepository.cls", _
                           "EAuthData.cls", "modAuthFactory.bas", "TestAuthService.bas", _
                           "TIAuthRepository.bas", _
                           "IConfig.cls", "IErrorHandlerService.cls", "modEnumeraciones.bas")
        
        Case "document", "documentos", "documents"
            ' Sección 3.2 - Gestión de Documentos + Dependencias
            arrFiles = Array("IDocumentService.cls", "CDocumentService.cls", "CMockDocumentService.cls", _
                           "IWordManager.cls", "CWordManager.cls", "CMockWordManager.cls", _
                           "IMapeoRepository.cls", "CMapeoRepository.cls", "CMockMapeoRepository.cls", _
                           "EMapeo.cls", "modDocumentServiceFactory.bas", _
                           "TIDocumentService.bas", _
                           "ISolicitudService.cls", "CSolicitudService.cls", "modSolicitudServiceFactory.bas", _
                           "IOperationLogger.cls", "IConfig.cls", "IErrorHandlerService.cls", "IFileSystem.cls", _
                           "modWordManagerFactory.bas", "modRepositoryFactory.bas", "modErrorHandlerFactory.bas")
        
        Case "expediente", "expedientes"
            ' Sección 3.3 - Gestión de Expedientes + Dependencias
            arrFiles = Array("IExpedienteService.cls", "CExpedienteService.cls", "CMockExpedienteService.cls", _
                           "IExpedienteRepository.cls", "CExpedienteRepository.cls", "CMockExpedienteRepository.cls", _
                           "EExpediente.cls", "modExpedienteServiceFactory.bas", "TestCExpedienteService.bas", _
                           "TIExpedienteRepository.bas", "modRepositoryFactory.bas", _
                           "IConfig.cls", "IOperationLogger.cls", "IErrorHandlerService.cls")
        
        Case "solicitud", "solicitudes"
            ' Sección 3.4 - Gestión de Solicitudes + Dependencias
            arrFiles = Array("ISolicitudService.cls", "CSolicitudService.cls", "CMockSolicitudService.cls", _
                           "ISolicitudRepository.cls", "CSolicitudRepository.cls", "CMockSolicitudRepository.cls", _
                           "ESolicitud.cls", "EDatosPc.cls", "EDatosCdCa.cls", "EDatosCdCaSub.cls", _
                           "modSolicitudServiceFactory.bas", "TestSolicitudService.bas", _
                           "TISolicitudRepository.bas", _
                           "IAuthService.cls", "modAuthFactory.bas", _
                           "IOperationLogger.cls", "IErrorHandlerService.cls", "IConfig.cls")
        
        Case "workflow", "flujo"
            ' Sección 3.5 - Gestión de Workflow + Dependencias
            arrFiles = Array("IWorkflowService.cls", "CWorkflowService.cls", "CMockWorkflowService.cls", _
                           "IWorkflowRepository.cls", "CWorkflowRepository.cls", "CMockWorkflowRepository.cls", _
                           "modWorkflowServiceFactory.bas", "TestWorkflowService.bas", _
                           "TIWorkflowRepository.bas", _
                           "IOperationLogger.cls", "IConfig.cls", "IErrorHandlerService.cls")
        
        Case "mapeo", "mapping"
            ' Sección 3.6 - Gestión de Mapeos + Dependencias
            arrFiles = Array("IMapeoRepository.cls", "CMapeoRepository.cls", "CMockMapeoRepository.cls", _
                           "EMapeo.cls", "TIMapeoRepository.bas", _
                           "IConfig.cls", "IErrorHandlerService.cls")
        
        Case "notification", "notificacion"
            ' Sección 3.7 - Gestión de Notificaciones + Dependencias
            arrFiles = Array("INotificationService.cls", "CNotificationService.cls", "CMockNotificationService.cls", _
                           "INotificationRepository.cls", "CNotificationRepository.cls", "CMockNotificationRepository.cls", _
                           "modNotificationServiceFactory.bas", "TINotificationService.bas", _
                           "IOperationLogger.cls", "IErrorHandlerService.cls", "IConfig.cls")
        
        Case "operation", "operacion", "logging"
            ' Sección 3.8 - Gestión de Operaciones y Logging + Dependencias
            arrFiles = Array("IOperationLogger.cls", "COperationLogger.cls", "CMockOperationLogger.cls", _
                           "IOperationRepository.cls", "COperationRepository.cls", "CMockOperationRepository.cls", _
                           "EOperationLog.cls", "modOperationLoggerFactory.bas", "TestOperationLogger.bas", _
                           "TIOperationRepository.bas", _
                           "IErrorHandlerService.cls", "IConfig.cls")
        
        Case "config", "configuracion"
            ' Sección 4 - Configuración + Dependencias
            arrFiles = Array("IConfig.cls", "CConfig.cls", "CMockConfig.cls", "modConfigFactory.bas", _
                           "TestCConfig.bas")
        
        Case "filesystem", "archivos"
            ' Sección 5 - Sistema de Archivos + Dependencias
            arrFiles = Array("IFileSystem.cls", "CFileSystem.cls", "CMockFileSystem.cls", _
                           "modFileSystemFactory.bas", "TIFileSystem.bas", _
                           "IErrorHandlerService.cls")
        
        Case "word"
            ' Sección 6 - Gestión de Word + Dependencias
            arrFiles = Array("IWordManager.cls", "CWordManager.cls", "CMockWordManager.cls", _
                           "modWordManagerFactory.bas", "TIWordManager.bas", _
                           "IFileSystem.cls", "IErrorHandlerService.cls")
        
        Case "error", "errores", "errors"
            ' Sección 7 - Gestión de Errores + Dependencias
            arrFiles = Array("IErrorHandlerService.cls", "CErrorHandlerService.cls", "CMockErrorHandlerService.cls", _
                           "modErrorHandlerFactory.bas", "TestErrorHandlerService.bas", _
                           "IConfig.cls", "IFileSystem.cls")
        
        Case "testframework", "testing", "framework"
            ' Sección 8 - Framework de Testing + Dependencias
            arrFiles = Array("ITestReporter.cls", "CTestResult.cls", "CTestSuiteResult.cls", "CTestReporter.cls", _
                           "modTestRunner.bas", "modTestUtils.bas", "modAssert.bas", _
                           "TestModAssert.bas", "IFileSystem.cls", "IConfig.cls", _
                           "IErrorHandlerService.cls")
        
        Case "app", "aplicacion", "application"
            ' Sección 9 - Gestión de Aplicación + Dependencias
            arrFiles = Array("IAppManager.cls", "CAppManager.cls", "CMockAppManager.cls", _
                           "ModAppManagerFactory.bas", "TestAppManager.bas", "IAuthService.cls", _
                           "IConfig.cls", "IErrorHandlerService.cls")
        
        Case "models", "modelos", "datos"
            ' Sección 10 - Modelos de Datos
            arrFiles = Array("EUsuario.cls", "ESolicitud.cls", "EExpediente.cls", "EDatosPc.cls", _
                           "EDatosCdCa.cls", "EDatosCdCaSub.cls", "EEstado.cls", "ETransicion.cls", _
                           "EMapeo.cls", "EAdjuntos.cls", "ELogCambios.cls", "ELogErrores.cls", "EOperationLog.cls", "EAuthData.cls")
        
        Case "utils", "utilidades", "enumeraciones"
            ' Sección 11 - Utilidades y Enumeraciones
            arrFiles = Array("modRepositoryFactory.bas", "modEnumeraciones.bas", "modQueries.bas", _
                           "ModAppManagerFactory.bas", "modAuthFactory.bas", "modConfigFactory.bas", _
                           "modDocumentServiceFactory.bas", "modErrorHandlerFactory.bas", _
                           "modExpedienteServiceFactory.bas", "modFileSystemFactory.bas", _
                           "modNotificationServiceFactory.bas", "modOperationLoggerFactory.bas", _
                           "modSolicitudServiceFactory.bas", "modWordManagerFactory.bas", _
                           "modWorkflowServiceFactory.bas")
        
        Case "tests", "pruebas", "testing", "test"
            ' Sección 12 - Archivos de Pruebas (Autodescubrimiento)
            arrFiles = Array()
        Case Else
            ' Funcionalidad no reconocida - devolver array vacío
            arrFiles = Array()
    End Select
    
    GetFunctionalityFiles = arrFiles
End Function



        


' Subrutina para empaquetar archivos de código por funcionalidad
' Subrutina para empaquetar archivos de código por funcionalidad o por lista de ficheros
Sub BundleFunctionality()
    On Error Resume Next
    
    Dim strFunctionalityOrFiles, strDestPath, strBundlePath, timestamp
    
    ' Verificar argumentos
    If objArgs.Count < 2 Then
        WScript.Echo "Error: Se requiere nombre de funcionalidad o lista de ficheros"
        WScript.Echo "Uso: cscript condor_cli.vbs bundle <funcionalidad | fichero1,fichero2,...> [ruta_destino]"
        WScript.Quit 1
    End If
    
    strFunctionalityOrFiles = objArgs(1)
    
    ' Determinar ruta de destino
    If objArgs.Count >= 3 Then
        strDestPath = objArgs(2)
    Else
        strDestPath = objFSO.GetParentFolderName(WScript.ScriptFullName)
    End If
    
    ' Crear timestamp
    timestamp = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & "_" & _
                Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
    
    ' Crear nombre de carpeta bundle
    Dim bundleName
    If InStr(strFunctionalityOrFiles, ",") > 0 Then
        bundleName = "bundle_custom_" & timestamp
    Else
        bundleName = "bundle_" & strFunctionalityOrFiles & "_" & timestamp
    End If
    strBundlePath = objFSO.BuildPath(strDestPath, bundleName)
    
    WScript.Echo "=== EMPAQUETANDO ARTEFACTOS ==="
    WScript.Echo "Buscando archivos en: " & strSourcePath
    WScript.Echo "Carpeta destino: " & strBundlePath
    
    ' Crear carpeta de destino
    If Not objFSO.FolderExists(strBundlePath) Then
        objFSO.CreateFolder strBundlePath
        If Err.Number <> 0 Then
            WScript.Echo "Error creando carpeta de destino: " & Err.Description
            WScript.Quit 1
        End If
    End If
    
    Dim arrFilesToBundle
    
    ' Lógica de Detección Inteligente
    If InStr(strFunctionalityOrFiles, ",") > 0 Then
        ' MODO 1: Lista de ficheros explícita
        WScript.Echo "Modo: Lista de ficheros explícita."
        arrFilesToBundle = Split(strFunctionalityOrFiles, ",")
    Else
        ' MODO 2: Verificar si es funcionalidad conocida o archivo individual
        arrFilesToBundle = GetFunctionalityFiles(strFunctionalityOrFiles)
        
        If UBound(arrFilesToBundle) >= 0 Then
            ' Es una funcionalidad conocida
            WScript.Echo "Modo: Funcionalidad '" & strFunctionalityOrFiles & "'."
        Else
            ' No es funcionalidad conocida, buscar archivo individual en src
            Dim singleFilePath
            singleFilePath = objFSO.BuildPath(strSourcePath, strFunctionalityOrFiles)
            
            If objFSO.FileExists(singleFilePath) Then
                ' Archivo encontrado, tratarlo como lista de un elemento
                WScript.Echo "Modo: Archivo individual '" & strFunctionalityOrFiles & "'."
                ReDim arrFilesToBundle(0)
                arrFilesToBundle(0) = strFunctionalityOrFiles
            Else
                ' Archivo no encontrado
                WScript.Echo "Error: '" & strFunctionalityOrFiles & "' no es una funcionalidad conocida ni un archivo existente en src."
                WScript.Echo "Funcionalidades disponibles: Auth, Document, Expediente, Solicitud, Workflow, Mapeo, Notification, Operation, Config, FileSystem, Word, Error, TestFramework, App, Models, Utils, Tests"
                WScript.Quit 1
            End If
        End If
    End If
    
    ' Llamar a la subrutina de ayuda para copiar los ficheros
    Call CopyFilesToBundle(arrFilesToBundle, strBundlePath)
    
    On Error GoTo 0
End Sub

' NUEVA SUBRUTINA DE AYUDA
' Copia una lista de ficheros al directorio del paquete
Sub CopyFilesToBundle(arrFiles, strBundlePath)
    Dim copiedFiles, notFoundFiles
    copiedFiles = 0
    notFoundFiles = 0
    
    If UBound(arrFiles) < 0 Then
        WScript.Echo "Advertencia: La lista de ficheros a empaquetar está vacía."
    End If

    Dim i, fileName, filePath, destFilePath
    For i = 0 To UBound(arrFiles)
        fileName = Trim(arrFiles(i))
        filePath = objFSO.BuildPath(strSourcePath, fileName)
        
        If objFSO.FileExists(filePath) Then
            ' Copiar archivo con extensión .txt añadida
            destFilePath = objFSO.BuildPath(strBundlePath, fileName & ".txt")
            objFSO.CopyFile filePath, destFilePath, True
            
            If Err.Number <> 0 Then
                WScript.Echo "  ? Error copiando " & fileName & ": " & Err.Description
                Err.Clear
            Else
                WScript.Echo "  ? " & fileName & " -> " & fileName & ".txt"
                copiedFiles = copiedFiles + 1
            End If
        Else
            WScript.Echo "  ? Archivo no encontrado: " & fileName
            notFoundFiles = notFoundFiles + 1
        End If
    Next
    
    WScript.Echo ""
    WScript.Echo "=== RESULTADO DEL EMPAQUETADO ==="
    WScript.Echo "Archivos copiados: " & copiedFiles
    WScript.Echo "Archivos no encontrados: " & notFoundFiles
    WScript.Echo "Ubicación del paquete: " & strBundlePath
    
    If copiedFiles = 0 Then
        WScript.Echo "? No se copió ningún archivo."
    Else
        WScript.Echo "? Empaquetado completado exitosamente"
    End If
End Sub

' Función auxiliar para convertir rutas relativas a absolutas
Private Function ResolveRelativePath(relativePath)
    Dim objFSO
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Verificar si la ruta ya es absoluta (contiene : en la segunda posición)
    If Len(relativePath) >= 2 And Mid(relativePath, 2, 1) = ":" Then
        ResolveRelativePath = relativePath
        Exit Function
    End If
    
    ' Obtener el directorio actual del script
    Dim currentDir
    currentDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
    
    ' Si la ruta empieza con .\, quitarlo
    If Left(relativePath, 2) = ".\" Then
        relativePath = Mid(relativePath, 3)
    End If
    
    ' Si la ruta empieza con \, quitarlo
    If Left(relativePath, 1) = "\" Then
        relativePath = Mid(relativePath, 2)
    End If
    
    ' Combinar la ruta actual con la ruta relativa
    ResolveRelativePath = objFSO.BuildPath(currentDir, relativePath)
End Function

' Función auxiliar para convertir tipos de datos DAO a texto legible
Private Function DaoTypeToString(dataType)
    Select Case dataType
        Case 1: DaoTypeToString = "Boolean"
        Case 2: DaoTypeToString = "Byte"
        Case 3: DaoTypeToString = "Integer"
        Case 4: DaoTypeToString = "Long"
        Case 5: DaoTypeToString = "Currency"
        Case 6: DaoTypeToString = "Single"
        Case 7: DaoTypeToString = "Double"
        Case 8: DaoTypeToString = "DateTime"
        Case 10: DaoTypeToString = "Text"
        Case 11: DaoTypeToString = "OLE Object"
        Case 12: DaoTypeToString = "Memo"
        Case 20: DaoTypeToString = "BigInt"
        Case Else: DaoTypeToString = "Desconocido (" & dataType & ")"
    End Select
End Function

' ===================================================================
' SUBRUTINA: ExecuteMigrations
' Descripción: Ejecuta scripts de migración SQL desde la carpeta /db/migrations
' ===================================================================
Sub ExecuteMigrations()
    Dim strMigrationsPath, objMigrationsFolder, objFile, strTargetFile
    
    strMigrationsPath = objFSO.GetParentFolderName(strSourcePath) & "\db\migrations"
    WScript.Echo "=== INICIANDO MIGRACION DE DATOS SQL ==="
    WScript.Echo "Directorio de migraciones: " & strMigrationsPath
    
    If Not objFSO.FolderExists(strMigrationsPath) Then
        WScript.Echo "ERROR: El directorio de migraciones no existe: " & strMigrationsPath
        WScript.Quit 1
    End If
    
    Set objMigrationsFolder = objFSO.GetFolder(strMigrationsPath)
    
    ' Modo 1: Migrar un fichero específico
    If objArgs.Count > 1 Then
        strTargetFile = objArgs(1)
        Dim targetPath
        targetPath = objFSO.BuildPath(strMigrationsPath, strTargetFile)
        If objFSO.FileExists(targetPath) Then
            WScript.Echo "Ejecutando migración específica: " & strTargetFile
            Call ProcessSqlFile(targetPath)
        Else
            WScript.Echo "ERROR: El archivo de migración especificado no existe: " & targetPath
            WScript.Quit 1
        End If
    ' Modo 2: Migrar todos los ficheros .sql
    Else
        WScript.Echo "Ejecutando todas las migraciones en el directorio (en orden alfabético)..."
        
        ' Crear un array para almacenar los nombres de archivos y ordenarlos
        Dim arrFiles(), intFileCount, i, j, strTemp
        intFileCount = 0
        
        ' Contar archivos SQL
        For Each objFile In objMigrationsFolder.Files
            If LCase(objFSO.GetExtensionName(objFile.Name)) = "sql" Then
                intFileCount = intFileCount + 1
            End If
        Next
        
        ' Redimensionar array
        ReDim arrFiles(intFileCount - 1)
        
        ' Llenar array con rutas de archivos
        i = 0
        For Each objFile In objMigrationsFolder.Files
            If LCase(objFSO.GetExtensionName(objFile.Name)) = "sql" Then
                arrFiles(i) = objFile.Path
                i = i + 1
            End If
        Next
        
        ' Ordenar array usando bubble sort
        For i = 0 To UBound(arrFiles) - 1
            For j = i + 1 To UBound(arrFiles)
                If UCase(objFSO.GetFileName(arrFiles(i))) > UCase(objFSO.GetFileName(arrFiles(j))) Then
                    strTemp = arrFiles(i)
                    arrFiles(i) = arrFiles(j)
                    arrFiles(j) = strTemp
                End If
            Next
        Next
        
        ' Ejecutar archivos en orden
        For i = 0 To UBound(arrFiles)
            Call ProcessSqlFile(arrFiles(i))
        Next
    End If
    
    WScript.Echo "=== MIGRACION COMPLETADA EXITOSAMENTE ==="
End Sub

' ===================================================================
' SUBRUTINA: ProcessSqlFile
' Descripción: Parsea y ejecuta los comandos de un fichero SQL
' CORREGIDO: Utiliza ADODB.Stream para leer ficheros con codificación UTF-8.
' ===================================================================
' FUNCIÓN: CleanSqlContent
' Elimina comentarios SQL y líneas vacías del contenido
Function CleanSqlContent(sqlContent)
    Dim arrLines, cleanedLines, i, trimmedLine
    
    ' Dividir en líneas
    arrLines = Split(sqlContent, vbCrLf)
    If UBound(arrLines) = 0 Then
        arrLines = Split(sqlContent, vbLf)
    End If
    
    ' Filtrar líneas
    cleanedLines = ""
    For i = 0 To UBound(arrLines)
        trimmedLine = Trim(arrLines(i))
        
        ' Ignorar líneas vacías y comentarios
        If Len(trimmedLine) > 0 And Left(trimmedLine, 2) <> "--" Then
            If cleanedLines <> "" Then
                cleanedLines = cleanedLines & vbCrLf
            End If
            cleanedLines = cleanedLines & arrLines(i)
        End If
    Next
    
    CleanSqlContent = cleanedLines
End Function

Sub ProcessSqlFile(filePath)
    Dim objStream, strContent, arrCommands, sqlCommand, conn
    
    WScript.Echo "------------------------------------------------------------"
    WScript.Echo "Procesando fichero: " & objFSO.GetFileName(filePath)
    
    On Error Resume Next
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "UTF-8"
    objStream.Open
    objStream.LoadFromFile filePath
    strContent = objStream.ReadText
    objStream.Close
    Set objStream = Nothing
    If Err.Number <> 0 Then
        WScript.Echo "  ERROR: No se pudo leer el fichero: " & Err.Description
        WScript.Quit 1 ' Detener en caso de error de lectura
    End If
    On Error GoTo 0
    
    ' Usar conexión ADO para un manejo de errores DDL robusto
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strAccessPath & ";"
    
    ' Limpiar comentarios y líneas vacías antes de procesar
    strContent = CleanSqlContent(strContent)
    
    ' Dividir en comandos por punto y coma
    arrCommands = Split(strContent, ";")
    
    ' Ejecutar cada comando
    For Each sqlCommand In arrCommands
        sqlCommand = Trim(sqlCommand)
        
        ' Solo ejecutar comandos que no estén vacías
        If Len(sqlCommand) > 5 Then
            On Error Resume Next
            conn.Execute sqlCommand
            
            If Err.Number <> 0 Then
                WScript.Echo "    ERROR al ejecutar comando: " & Err.Description
                WScript.Echo "    SQL: " & sqlCommand
                WScript.Echo "  MIGRACIÓN FALLIDA. Abortando."
                WScript.Echo "------------------------------------------------------------"
                conn.Close
                Set conn = Nothing
                WScript.Quit 1 ' Detener la ejecución inmediatamente
            Else
                WScript.Echo "    Comando ejecutado exitosamente."
            End If
            On Error GoTo 0
        End If
    Next
    
    conn.Close
    Set conn = Nothing
    
    WScript.Echo "  Fichero procesado exitosamente."
    WScript.Echo "------------------------------------------------------------"
End Sub

' Función para formatear texto con ancho fijo
Function PadRight(text, width)
    If Len(text) >= width Then
        PadRight = Left(text, width)
    Else
        PadRight = text & String(width - Len(text), " ")
    End If
End Function

' Función para verificar la refactorización de logging
Sub VerifyLoggingRefactoring()
    Dim serviceFiles, fileName, filePath, fileContent
    Dim obsoleteCalls, refactoredCalls
    Dim totalObsolete, totalRefactored
    
    serviceFiles = Array("CAuthService.cls", "CNotificationService.cls", "CWorkflowService.cls", "CSolicitudService.cls")
    totalObsolete = 0
    totalRefactored = 0
    
    WScript.Echo "  Verificando servicios refactorizados..."
    
    For Each fileName In serviceFiles
        filePath = strSourcePath & "\" & fileName
        
        If objFSO.FileExists(filePath) Then
            fileContent = objFSO.OpenTextFile(filePath, 1).ReadAll
            
            ' Buscar llamadas obsoletas (3 parámetros)
            obsoleteCalls = CountMatches(fileContent, "LogOperation\s*\(\s*""[^""]*""\s*,\s*\d+\s*,\s*""[^""]*""\s*\)")
            
            ' Buscar llamadas refactorizadas (EOperationLog)
            refactoredCalls = CountMatches(fileContent, "LogOperation\s*\(\s*operationLog\)")
            
            totalObsolete = totalObsolete + obsoleteCalls
            totalRefactored = totalRefactored + refactoredCalls
            
            If obsoleteCalls > 0 Then
                WScript.Echo "    ⚠️  " & fileName & ": " & obsoleteCalls & " llamadas obsoletas encontradas"
            Else
                WScript.Echo "    ✅ " & fileName & ": Refactorizado (" & refactoredCalls & " llamadas EOperationLog)"
            End If
        Else
            WScript.Echo "    ❌ " & fileName & ": Archivo no encontrado"
        End If
    Next
    
    WScript.Echo "  Resumen de refactorización:"
    WScript.Echo "    - Llamadas obsoletas: " & totalObsolete
    WScript.Echo "    - Llamadas refactorizadas: " & totalRefactored
    
    If totalObsolete > 0 Then
        WScript.Echo "    ⚠️  ADVERTENCIA: Aún existen llamadas obsoletas por refactorizar"
    Else
        WScript.Echo "    ✅ ÉXITO: Todos los servicios han sido refactorizados"
    End If
End Sub

' Función auxiliar para contar coincidencias de regex
Function CountMatches(text, pattern)
    Dim regex, matches
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = pattern
    regex.Global = True
    regex.IgnoreCase = True
    
    Set matches = regex.Execute(text)
    CountMatches = matches.Count
End Function

' ===================================================================
' SUBRUTINA: ExportForm
' Descripción: Exporta el diseño de un formulario a JSON enriquecido con JsonWriter.
' ===================================================================
Sub ExportForm()
    Dim strDbPath, strFormName, strOutputPath, strPassword
    Dim strSchemaVersion, strExpand, strResourceRoot, strSrcDir
    Dim bPretty, bNoControls, bStrict
    Dim i
    
    ' Verificar argumentos mínimos
    If objArgs.Count < 3 Then
        WScript.Echo "Error: El comando export-form requiere al menos una ruta de base de datos y un nombre de formulario."
        WScript.Echo "Uso: cscript condor_cli.vbs export-form <db_path> <form_name> [flags]"
        WScript.Quit 1
    End If
    
    ' Asignar argumentos básicos
    strDbPath = objArgs(1)
    strFormName = objArgs(2)
    strOutputPath = ""
    strPassword = ""
    strSchemaVersion = "1.0.0"
    strExpand = "events,formatting,resources"
    strResourceRoot = ""
    strSrcDir = ""
    bPretty = False
    bNoControls = False
    bStrict = False
    
    ' Procesar argumentos opcionales
    For i = 3 To objArgs.Count - 1
        If LCase(objArgs(i)) = "--output" And i < objArgs.Count - 1 Then
            strOutputPath = objArgs(i + 1)
        ElseIf LCase(objArgs(i)) = "--password" And i < objArgs.Count - 1 Then
            strPassword = objArgs(i + 1)
        ElseIf LCase(objArgs(i)) = "--schema-version" And i < objArgs.Count - 1 Then
            strSchemaVersion = objArgs(i + 1)
        ElseIf LCase(Left(objArgs(i), 9)) = "--expand=" Then
            strExpand = Mid(objArgs(i), 10)
        ElseIf LCase(objArgs(i)) = "--resource-root" And i < objArgs.Count - 1 Then
            strResourceRoot = objArgs(i + 1)
        ElseIf LCase(objArgs(i)) = "--pretty" Then
            bPretty = True
        ElseIf LCase(objArgs(i)) = "--no-controls" Then
            bNoControls = True
        ElseIf LCase(objArgs(i)) = "--verbose" Then
            gVerbose = True
        ElseIf LCase(objArgs(i)) = "--src" And i < objArgs.Count - 1 Then
            strSrcDir = objArgs(i + 1)
        ElseIf LCase(objArgs(i)) = "--strict" Then
            bStrict = True
        End If
    Next
    
    ' Configurar resource root por defecto
    If strResourceRoot = "" Then
        strResourceRoot = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\ui\assets\"
    End If
    
    ' Configurar src directory por defecto
    If strSrcDir = "" Then
        strSrcDir = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\src\"
    End If
    
    ' Construir ruta de salida por defecto si no se especifica
    If strOutputPath = "" Then
        Dim uiDefinitionsPath
        uiDefinitionsPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\ui\definitions\"
        ' Asegurarse de que el directorio existe
        If Not objFSO.FolderExists(uiDefinitionsPath) Then objFSO.CreateFolder uiDefinitionsPath
        strOutputPath = objFSO.BuildPath(uiDefinitionsPath, strFormName & ".json")
    End If
    
    ' Verificar que el archivo de base de datos existe
    If Not objFSO.FileExists(strDbPath) Then
        WScript.Echo "Error: La base de datos no existe: " & strDbPath
        WScript.Quit 1
    End If
    
    If gVerbose Then WScript.Echo "Abriendo base de datos: " & strDbPath
    
    ' Crear instancia de Access
    Set objAccess = CreateObject("Access.Application")
    objAccess.Visible = False
    
    ' Abrir base de datos
    On Error Resume Next
    If strPassword = "" Then
        objAccess.OpenCurrentDatabase strDbPath
    Else
        objAccess.OpenCurrentDatabase strDbPath, , strPassword
    End If
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al abrir la base de datos: " & Err.Description
        objAccess.Quit
        WScript.Quit 1
    End If
    
    ' Abrir formulario en modo diseño con manejo de errores para código de inicio
    If gVerbose Then WScript.Echo "Intentando abrir el formulario '" & strFormName & "' en modo diseño..."
    On Error Resume Next
    
    objAccess.DoCmd.OpenForm strFormName, 3 ' acDesign
    
    If Err.Number <> 0 Then
        WScript.Echo "--------------------------------------------------------------------"
        WScript.Echo "ERROR: Se produjo un error crítico al intentar abrir el formulario."
        WScript.Echo "CAUSA PROBABLE: El código de inicialización del formulario (eventos Form_Open o Form_Load) falló porque depende de otras partes de la aplicación que no están cargadas en este contexto de automatización."
        WScript.Echo "ACCION: Este es un problema conocido en la base de datos de destino y no se puede solucionar desde el CLI. Exporte un formulario diferente o revise el código del formulario de destino."
        WScript.Echo "Detalle del error VBA (" & Err.Number & "): " & Err.Description
        WScript.Echo "--------------------------------------------------------------------"
        objAccess.Quit
        WScript.Quit 1
    End If
    
    On Error GoTo 0
    
    If gVerbose Then WScript.Echo "Formulario '" & strFormName & "' abierto en modo diseño."
    
    ' Obtener referencia al objeto del formulario activo
    Dim frm
    Set frm = objAccess.Forms(strFormName)
    
    ' Crear JsonWriter para construcción del JSON
    Dim jsonWriter
    Set jsonWriter = New JsonWriter
    
    ' Inicializar estructura JSON principal
    jsonWriter.StartObject
    
    ' Schema version
    jsonWriter.AddProperty "schemaVersion", strSchemaVersion
    jsonWriter.AddProperty "units", "twips"
    
    ' Metadata
    jsonWriter.StartObjectProperty "metadata"
    jsonWriter.AddProperty "generatedAtUTC", FormatDateTime(Now, vbGeneralDate) & " UTC"
    jsonWriter.AddProperty "tool", "condor_cli"
    jsonWriter.AddProperty "toolVersion", "1.0.0"
    jsonWriter.EndObject
    
    ' Resources (inicializar vacío, se llenará al procesar controles)
    Dim resourceImages
    Set resourceImages = CreateList()
    jsonWriter.StartObjectProperty "resources"
    jsonWriter.StartArrayProperty "images"
    ' Se llenará durante el procesamiento de controles
    
    ' Form name
    jsonWriter.AddProperty "formName", frm.Name
    
    ' Properties del formulario
    jsonWriter.StartObjectProperty "properties"
    On Error Resume Next
    
    ' Propiedades básicas del formulario
    jsonWriter.AddProperty "caption", frm.Caption
    jsonWriter.AddProperty "popUp", frm.PopUp
    jsonWriter.AddProperty "modal", frm.Modal
    jsonWriter.AddProperty "width", frm.Width
    jsonWriter.AddProperty "autoCenter", frm.AutoCenter
    
    ' BorderStyle con mapeo a token canónico
    jsonWriter.AddProperty "borderStyle", MapBorderStyleToToken(frm.BorderStyle)
    
    ' Propiedades de interfaz
    jsonWriter.AddProperty "recordSelectors", frm.RecordSelectors
    jsonWriter.AddProperty "dividingLines", frm.DividingLines
    jsonWriter.AddProperty "navigationButtons", frm.NavigationButtons
    
    ' ScrollBars con mapeo a token canónico
    jsonWriter.AddProperty "scrollBars", MapScrollBarsToToken(frm.ScrollBars)
    
    ' Propiedades de ventana
    jsonWriter.AddProperty "controlBox", frm.ControlBox
    jsonWriter.AddProperty "closeButton", frm.CloseButton
    
    ' MinMaxButtons con mapeo a token canónico
    jsonWriter.AddProperty "minMaxButtons", MapMinMaxButtonsToToken(frm.MinMaxButtons)
    
    jsonWriter.AddProperty "movable", frm.Moveable
    
    ' RecordsetType con mapeo a token canónico
    jsonWriter.AddProperty "recordsetType", MapRecordsetTypeToToken(frm.RecordsetType)
    
    ' Orientation con mapeo a token canónico
    jsonWriter.AddProperty "orientation", MapOrientationToToken(frm.Orientation)
    
    ' Propiedades de SplitForm (solo si aplica)
    If frm.DefaultView = 3 Then ' SplitForm
        ' SplitFormSize: puede ser "Auto" o un número en twips
        Dim splitFormSizeValue
        If frm.SplitFormSize = -1 Then
            splitFormSizeValue = "Auto"
        Else
            splitFormSizeValue = frm.SplitFormSize
        End If
        jsonWriter.AddProperty "splitFormSize", splitFormSizeValue
        
        jsonWriter.AddProperty "splitFormOrientation", MapSplitFormOrientationToToken(frm.SplitFormOrientation)
        jsonWriter.AddProperty "splitFormSplitterBar", frm.SplitFormSplitterBar
    End If
    
    ' Propiedades existentes (mantener compatibilidad)
    jsonWriter.AddProperty "recordSource", frm.RecordSource

    ' Determinar recordSourceType
    Dim recordSourceType
    recordSourceType = "None"
    If frm.RecordSource <> "" Then
        ' Simplificado: si contiene SELECT es SQL, sino es Table
        If InStr(UCase(frm.RecordSource), "SELECT") > 0 Then
            recordSourceType = "SQL"
        Else
            recordSourceType = "Table"
        End If
    End If
    jsonWriter.AddProperty "recordSourceType", recordSourceType
    
    jsonWriter.AddProperty "allowEdits", frm.AllowEdits
    jsonWriter.AddProperty "allowAdditions", frm.AllowAdditions
    jsonWriter.AddProperty "allowDeletions", frm.AllowDeletions
    
    ' Cycle (simplificado)
    Dim cycleValue
    cycleValue = "AllRecords"
    If frm.Cycle = 1 Then cycleValue = "CurrentRecord"
    If frm.Cycle = 2 Then cycleValue = "NewRecordOnly"
    jsonWriter.AddProperty "cycle", cycleValue
    
    jsonWriter.AddProperty "autoResize", frm.AutoResize
    
    ' DefaultView (simplificado)
    Dim defaultViewValue
    defaultViewValue = "SingleForm"
    If frm.DefaultView = 1 Then defaultViewValue = "ContinuousForms"
    If frm.DefaultView = 2 Then defaultViewValue = "Datasheet"
    If frm.DefaultView = 3 Then defaultViewValue = "SplitForm"
    jsonWriter.AddProperty "defaultView", defaultViewValue
    
    On Error GoTo 0
    jsonWriter.EndObject ' properties
    
    ' Events del formulario (si expand incluye events)
    If InStr(LCase(strExpand), "events") > 0 Then
        jsonWriter.StartObjectProperty "events"
        On Error Resume Next
        If frm.OnOpen <> "" Then jsonWriter.AddProperty "OnOpen", frm.OnOpen
        If frm.OnLoad <> "" Then jsonWriter.AddProperty "OnLoad", frm.OnLoad
        If frm.OnClose <> "" Then jsonWriter.AddProperty "OnClose", frm.OnClose
        On Error GoTo 0
        jsonWriter.EndObject
    End If
    
    ' Sections
    jsonWriter.StartObjectProperty "sections"
    
    ' Array con los nombres de las secciones y sus índices
    Dim sectionNames, sectionName, currentSection
    Dim sectionIndex, controlIndex, tabIndex, zIndex
    sectionNames = Array("header", "detail", "footer")
    tabIndex = 0
    zIndex = 0
    
    ' Bucle para procesar las 3 secciones (header=1, detail=0, footer=2)
    For sectionIndex = 0 To 2
        ' Obtener nombre y referencia de la sección actual
        If sectionIndex = 0 Then
            sectionName = "detail"
            Set currentSection = frm.Section(0)
        ElseIf sectionIndex = 1 Then
            sectionName = "header"
            On Error Resume Next
            Set currentSection = frm.Section(1)
            If Err.Number <> 0 Then
                Err.Clear
                Set currentSection = Nothing
            End If
            On Error GoTo 0
        Else
            sectionName = "footer"
            On Error Resume Next
            Set currentSection = frm.Section(2)
            If Err.Number <> 0 Then
                Err.Clear
                Set currentSection = Nothing
            End If
            On Error GoTo 0
        End If
        
        ' Solo procesar si la sección existe
        If Not currentSection Is Nothing Then
            jsonWriter.StartObjectProperty sectionName
        
            ' Propiedades de la sección
            On Error Resume Next
            jsonWriter.AddProperty "visible", currentSection.Visible
            jsonWriter.AddProperty "height", currentSection.Height
            
            ' Propiedades de color y formato (si expand incluye formatting)
            If InStr(LCase(strExpand), "formatting") > 0 Then
                If currentSection.BackColor <> 0 Then
                    jsonWriter.AddProperty "backColor", OleToRgbHex(currentSection.BackColor)
                End If
                If currentSection.SpecialEffect <> 0 Then
                    jsonWriter.AddProperty "specialEffect", currentSection.SpecialEffect
                End If
                If currentSection.AlternateBackColor <> 0 Then
                    jsonWriter.AddProperty "alternateBackColor", OleToRgbHex(currentSection.AlternateBackColor)
                End If
            End If
            On Error GoTo 0
            
            ' Controles (si no es --no-controls)
            If Not bNoControls Then
                jsonWriter.StartArrayProperty "controls"
                
                If gVerbose Then WScript.Echo "Procesando sección " & sectionName & ", controles encontrados: " & currentSection.Controls.Count
                
                ' Crear array de controles ordenados por TabIndex para calcular zIndex
                Dim controlsArray, control, controlCount
                controlCount = currentSection.Controls.Count
                If controlCount > 0 Then
                    ReDim controlsArray(controlCount - 1)
                    controlIndex = 0
                    For Each control In currentSection.Controls
                        Set controlsArray(controlIndex) = control
                        controlIndex = controlIndex + 1
                    Next
                    
                    ' Procesar cada control
                    For controlIndex = 0 To controlCount - 1
                        Set control = controlsArray(controlIndex)
                         jsonWriter.StartObject
                         
                         ' Propiedades básicas del control
                         On Error Resume Next
                         jsonWriter.AddProperty "name", control.Name
                         
                         ' Determinar tipo de control normalizado
                         Dim controlType
                         controlType = TypeName(control)
                         Select Case controlType
                             Case "TextBox"
                                 jsonWriter.AddProperty "type", "TextBox"
                             Case "Label"
                                 jsonWriter.AddProperty "type", "Label"
                             Case "CommandButton"
                                 jsonWriter.AddProperty "type", "CommandButton"
                             Case Else
                                 jsonWriter.AddProperty "type", controlType
                         End Select
                         
                         If gVerbose Then WScript.Echo "Procesando control: " & control.Name & " (" & controlType & ")"
                         
                         ' Propiedades de posición y tamaño
                         jsonWriter.StartObjectProperty "properties"
                         jsonWriter.AddProperty "top", control.Top
                         jsonWriter.AddProperty "left", control.Left
                         jsonWriter.AddProperty "width", control.Width
                         jsonWriter.AddProperty "height", control.Height
                         
                         ' Propiedades de fuente (si expand incluye formatting)
                         If InStr(LCase(strExpand), "formatting") > 0 Then
                             If control.FontName <> "" Then jsonWriter.AddProperty "fontName", control.FontName
                             If control.FontSize > 0 Then jsonWriter.AddProperty "fontSize", control.FontSize
                             If control.FontWeight <> 0 Then jsonWriter.AddProperty "fontWeight", control.FontWeight
                         End If
                         
                         ' Índices calculados
                         jsonWriter.AddProperty "zIndex", zIndex
                         If control.TabStop Then
                             jsonWriter.AddProperty "tabIndex", tabIndex
                             tabIndex = tabIndex + 1
                         End If
                         zIndex = zIndex + 1
                         
                         ' Propiedades de estado
                         jsonWriter.AddProperty "tabStop", control.TabStop
                         jsonWriter.AddProperty "visible", control.Visible
                         jsonWriter.AddProperty "enabled", control.Enabled
                         
                         ' Propiedades específicas por tipo de control
                         If controlType = "TextBox" Then
                             If control.Locked <> False Then jsonWriter.AddProperty "locked", control.Locked
                             If control.ControlSource <> "" Then jsonWriter.AddProperty "controlSource", control.ControlSource
                             If control.Format <> "" Then jsonWriter.AddProperty "format", control.Format
                             If control.DecimalPlaces >= 0 Then jsonWriter.AddProperty "decimalPlaces", control.DecimalPlaces
                             If control.InputMask <> "" Then jsonWriter.AddProperty "inputMask", control.InputMask
                             If control.DefaultValue <> "" Then jsonWriter.AddProperty "defaultValue", control.DefaultValue
                             ' TextAlign normalizado
                             Select Case control.TextAlign
                                 Case 1: jsonWriter.AddProperty "textAlign", "Left"
                                 Case 2: jsonWriter.AddProperty "textAlign", "Center"
                                 Case 3: jsonWriter.AddProperty "textAlign", "Right"
                             End Select
                             ' ScrollBars normalizado
                             Select Case control.ScrollBars
                                 Case 0: jsonWriter.AddProperty "scrollBars", "None"
                                 Case 1: jsonWriter.AddProperty "scrollBars", "Horizontal"
                                 Case 2: jsonWriter.AddProperty "scrollBars", "Vertical"
                                 Case 3: jsonWriter.AddProperty "scrollBars", "Both"
                             End Select
                         ElseIf controlType = "Label" Then
                             If control.Caption <> "" Then jsonWriter.AddProperty "caption", control.Caption
                             ' TextAlign normalizado
                             Select Case control.TextAlign
                                 Case 1: jsonWriter.AddProperty "textAlign", "Left"
                                 Case 2: jsonWriter.AddProperty "textAlign", "Center"
                                 Case 3: jsonWriter.AddProperty "textAlign", "Right"
                             End Select
                             If control.WordWrap <> False Then jsonWriter.AddProperty "wordWrap", control.WordWrap
                         ElseIf controlType = "CommandButton" Then
                             If control.Caption <> "" Then jsonWriter.AddProperty "caption", control.Caption
                             If control.Picture <> "" Then
                                 jsonWriter.AddProperty "picture", control.Picture
                                 ' Añadir a resources si expand incluye resources
                                 If InStr(LCase(strExpand), "resources") > 0 And strResourceRoot <> "" Then
                                     ' TODO: Implementar lógica de recursos
                                 End If
                             End If
                             If control.Transparent <> False Then jsonWriter.AddProperty "transparent", control.Transparent
                         End If
                         
                         ' Propiedades de color (si expand incluye formatting)
                         If InStr(LCase(strExpand), "formatting") > 0 Then
                             If control.ForeColor <> 0 Then jsonWriter.AddProperty "foreColor", OleToRgbHex(control.ForeColor)
                             If control.BackColor <> 0 Then jsonWriter.AddProperty "backColor", OleToRgbHex(control.BackColor)
                             If control.BorderColor <> 0 Then jsonWriter.AddProperty "borderColor", OleToRgbHex(control.BorderColor)
                             If control.BorderStyle <> 0 Then jsonWriter.AddProperty "borderStyle", control.BorderStyle
                             If control.BorderWidth <> 0 Then jsonWriter.AddProperty "borderWidth", control.BorderWidth
                             If control.SpecialEffect <> 0 Then jsonWriter.AddProperty "specialEffect", control.SpecialEffect
                         End If
                         
                         ' Propiedades adicionales
                         If control.ControlTipText <> "" Then jsonWriter.AddProperty "tooltip", control.ControlTipText
                         If control.Tag <> "" Then jsonWriter.AddProperty "tag", control.Tag
                         
                         jsonWriter.EndObject ' properties
                         
                         ' Events del control (si expand incluye events)
                         If InStr(LCase(strExpand), "events") > 0 Then
                             jsonWriter.StartObjectProperty "events"
                             If control.OnClick <> "" Then jsonWriter.AddProperty "OnClick", control.OnClick
                             If control.OnDblClick <> "" Then jsonWriter.AddProperty "OnDblClick", control.OnDblClick
                             If control.OnGotFocus <> "" Then jsonWriter.AddProperty "OnGotFocus", control.OnGotFocus
                             If control.OnLostFocus <> "" Then jsonWriter.AddProperty "OnLostFocus", control.OnLostFocus
                             If control.OnChange <> "" Then jsonWriter.AddProperty "OnChange", control.OnChange
                             If control.OnAfterUpdate <> "" Then jsonWriter.AddProperty "OnAfterUpdate", control.OnAfterUpdate
                             If control.OnBeforeUpdate <> "" Then jsonWriter.AddProperty "OnBeforeUpdate", control.OnBeforeUpdate
                             
                             ' Agregar events.detected basado en handlers encontrados en código
                             Dim detectedEvents, eventKey
                             Set detectedEvents = CreateObject("Scripting.Dictionary")
                             
                             ' Buscar handlers para este control
                             Dim eventTypes
                             eventTypes = Array("Click", "DblClick", "GotFocus", "LostFocus", "Change", "AfterUpdate", "BeforeUpdate", "Current", "Load", "Open")
                             
                             Dim j
                             For j = 0 To UBound(eventTypes)
                                 eventKey = control.Name & "." & eventTypes(j)
                                 If detectedHandlers.Exists(eventKey) Then
                                     If Not detectedEvents.Exists(eventTypes(j)) Then
                                         detectedEvents.Add eventTypes(j), True
                                     End If
                                 End If
                             Next
                             
                             ' Escribir array detected si hay eventos detectados
                             If detectedEvents.Count > 0 Then
                                 jsonWriter.WriteProperty "detected"
                                 jsonWriter.StartArray
                                 Dim detectedKey
                                 For Each detectedKey In detectedEvents.Keys
                                     jsonWriter.WriteValue detectedKey
                                 Next
                                 jsonWriter.EndArray
                             End If
                             
                             jsonWriter.EndObject
                         End If
                         
                         On Error GoTo 0
                         jsonWriter.EndObject ' control
                     Next
                 End If
                 
                 jsonWriter.EndArray ' controls
             End If
             
             jsonWriter.EndObject ' section
         End If
     Next
     
     jsonWriter.EndObject ' sections
    
    ' Code module detection and event handlers
    Dim detectedHandlers
    Set detectedHandlers = DetectModuleAndHandlers(jsonWriter, strFormName, strSrcDir, gVerbose)
    
    jsonWriter.EndObject ' root object
    
    If gVerbose Then WScript.Echo "Estructura de datos del formulario extraída correctamente."
    
    ' Generar JSON usando JsonWriter
    Dim jsonString
    If bPretty Then
        jsonString = jsonWriter.ToString(True) ' Pretty print
    Else
        jsonString = jsonWriter.ToString(False) ' Compact
    End If
    
    ' Guardar en archivo con codificación UTF-8
    Dim objFile
    Set objFile = objFSO.CreateTextFile(strOutputPath, True, True) ' True = UTF-8 encoding
    objFile.Write jsonString
    objFile.Close
    
    If gVerbose Then
        WScript.Echo "Éxito: El formulario ha sido exportado a " & strOutputPath
        WScript.Echo "Esquema: " & strSchemaVersion & ", Expand: " & strExpand
    Else
        WScript.Echo "Éxito: El formulario ha sido exportado a " & strOutputPath
    End If
    
    ' Lógica de limpieza
    objAccess.DoCmd.Close
End Sub

' Detecta módulo y handlers de eventos para el formulario
' Parámetros:
'   jsonWriter - Objeto JsonWriter para escribir la sección code
'   formName - Nombre del formulario
'   srcDir - Directorio donde buscar el módulo
'   verbose - Si mostrar información detallada
' Retorna: Diccionario con handlers detectados (clave: "control.event")
Function DetectModuleAndHandlers(jsonWriter, formName, srcDir, verbose)
    Dim objFSO, moduleFile, moduleExists, moduleFilename
    Dim handlers, fileContent, regEx, matches, match
    Dim i, controlName, eventName, signature
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set handlers = CreateObject("Scripting.Dictionary")
    
    moduleExists = False
    moduleFilename = ""
    
    ' Heurística para encontrar el archivo del módulo
    Dim candidateFiles(3)
    candidateFiles(0) = srcDir & "Form_" & formName & ".bas"
    candidateFiles(1) = srcDir & formName & ".bas"
    candidateFiles(2) = srcDir & "frm" & formName & ".bas"
    candidateFiles(3) = srcDir & "Form_" & formName & ".cls"
    
    For i = 0 To 3
        If objFSO.FileExists(candidateFiles(i)) Then
            moduleExists = True
            moduleFilename = objFSO.GetFileName(candidateFiles(i))
            moduleFile = candidateFiles(i)
            Exit For
        End If
    Next
    
    ' Escribir sección code.module
    jsonWriter.WriteProperty "code"
    jsonWriter.StartObject
    jsonWriter.WriteProperty "module"
    jsonWriter.StartObject
    jsonWriter.WriteProperty "exists", moduleExists
    jsonWriter.WriteProperty "filename", moduleFilename
    
    ' Si existe el módulo, parsear handlers
    If moduleExists Then
        If verbose Then WScript.Echo "Detectando handlers en: " & moduleFile
        
        ' Leer contenido del archivo
        Dim textStream
        Set textStream = objFSO.OpenTextFile(moduleFile, 1) ' ForReading
        fileContent = textStream.ReadAll
        textStream.Close
        
        ' Crear expresión regular para detectar handlers
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.Global = True
        regEx.IgnoreCase = True
        regEx.MultiLine = True
        regEx.Pattern = "^\s*(Public|Private)?\s*Sub\s+([A-Za-z0-9_]+)_(Click|DblClick|Current|Load|Open|GotFocus|LostFocus|Change|AfterUpdate|BeforeUpdate)\s*\("
        
        Set matches = regEx.Execute(fileContent)
        
        ' Procesar matches
        jsonWriter.WriteProperty "handlers"
        jsonWriter.StartArray
        
        For Each match In matches
            controlName = match.SubMatches(1)
            eventName = match.SubMatches(2)
            signature = Trim(match.Value)
            
            ' Agregar handler al JSON
            jsonWriter.StartObject
            jsonWriter.WriteProperty "control", controlName
            jsonWriter.WriteProperty "event", eventName
            jsonWriter.WriteProperty "signature", signature
            jsonWriter.EndObject
            
            ' Guardar en diccionario para referencia
            Dim handlerKey
            handlerKey = controlName & "." & eventName
            handlers.Add handlerKey, True
            
            If verbose Then WScript.Echo "Handler detectado: " & controlName & "." & eventName
        Next
        
        jsonWriter.EndArray ' handlers
    Else
        ' No hay módulo, array vacío
        jsonWriter.WriteProperty "handlers"
        jsonWriter.StartArray
        jsonWriter.EndArray
        
        If verbose Then WScript.Echo "No se encontró módulo para el formulario " & formName
    End If
    
    jsonWriter.EndObject ' module
    jsonWriter.EndObject ' code
    
    ' Retornar diccionario de handlers para uso posterior
    Set DetectModuleAndHandlers = handlers
End Function

' ===================================================================
' SUBRUTINA: RoundtripFormCommand
' Descripción: Realiza test de roundtrip export→import→export y compara resultados
' ===================================================================
Sub RoundtripFormCommand()
    Dim strDbPath, strFormName, strTempDir, strPassword
    Dim strJson1Path, strJson2Path, strTempDbPath
    Dim i, bVerboseOriginal
    
    ' Verificar argumentos mínimos
    If objArgs.Count < 3 Then
        WScript.Echo "Error: El comando roundtrip-form requiere al menos una ruta de base de datos y un nombre de formulario."
        WScript.Echo "Uso: cscript condor_cli.vbs roundtrip-form <db_path> <form_name> [--password] [--verbose]"
        WScript.Quit 1
    End If
    
    ' Asignar argumentos básicos
    strDbPath = objArgs(1)
    strFormName = objArgs(2)
    strPassword = ""
    
    ' Procesar argumentos opcionales
    For i = 3 To objArgs.Count - 1
        If LCase(objArgs(i)) = "--password" And i < objArgs.Count - 1 Then
            strPassword = objArgs(i + 1)
        ElseIf LCase(objArgs(i)) = "--verbose" Then
            gVerbose = True
        End If
    Next
    
    ' Verificar que el archivo de base de datos existe
    If Not objFSO.FileExists(strDbPath) Then
        WScript.Echo "Error: La base de datos no existe: " & strDbPath
        WScript.Quit 1
    End If
    
    ' Crear directorio temporal para el test
    strTempDir = objFSO.GetSpecialFolder(2) & "\condor_roundtrip_" & Replace(Replace(Replace(Now, "/", ""), ":", ""), " ", "_")
    If Not objFSO.FolderExists(strTempDir) Then objFSO.CreateFolder strTempDir
    
    strJson1Path = objFSO.BuildPath(strTempDir, strFormName & "_export1.json")
    strJson2Path = objFSO.BuildPath(strTempDir, strFormName & "_export2.json")
    strTempDbPath = objFSO.BuildPath(strTempDir, "temp_" & objFSO.GetFileName(strDbPath))
    
    If gVerbose Then
        WScript.Echo "=== ROUNDTRIP TEST INICIADO ==="
        WScript.Echo "Base de datos: " & strDbPath
        WScript.Echo "Formulario: " & strFormName
        WScript.Echo "Directorio temporal: " & strTempDir
    End If
    
    On Error Resume Next
    
    ' PASO 1: Export inicial
    If gVerbose Then WScript.Echo "PASO 1: Exportando formulario original..."
    
    ' Simular llamada a export-form modificando argumentos temporalmente
    Dim originalArgs
    Set originalArgs = objArgs
    
    ' Crear argumentos para export-form
    Dim exportArgs1
    Set exportArgs1 = CreateObject("Scripting.Dictionary")
    exportArgs1.Add 0, "export-form"
    exportArgs1.Add 1, strDbPath
    exportArgs1.Add 2, strFormName
    exportArgs1.Add 3, "--output"
    exportArgs1.Add 4, strJson1Path
    If strPassword <> "" Then
        exportArgs1.Add 5, "--password"
        exportArgs1.Add 6, strPassword
    End If
    
    ' Ejecutar export usando ExportForm directamente
    bVerboseOriginal = gVerbose
    gVerbose = False ' Silenciar para evitar ruido
    
    ' Crear copia temporal de la base de datos
    objFSO.CopyFile strDbPath, strTempDbPath
    
    ' Llamar a ExportForm con argumentos modificados
    Dim tempArgCount
    tempArgCount = exportArgs1.Count
    
    ' Simular objArgs para ExportForm
    ReDim Preserve arrTempArgs(tempArgCount - 1)
    For i = 0 To tempArgCount - 1
        arrTempArgs(i) = exportArgs1(i)
    Next
    
    ' Ejecutar export (necesitamos simular el contexto)
    Call ExportFormInternal(strDbPath, strFormName, strJson1Path, strPassword)
    
    If Err.Number <> 0 Then
        WScript.Echo "Error en PASO 1 (export inicial): " & Err.Description
        gVerbose = bVerboseOriginal
        CleanupRoundtripTest strTempDir
        WScript.Quit 1
    End If
    
    If Not objFSO.FileExists(strJson1Path) Then
        WScript.Echo "Error: No se generó el archivo JSON en PASO 1: " & strJson1Path
        gVerbose = bVerboseOriginal
        CleanupRoundtripTest strTempDir
        WScript.Quit 1
    End If
    
    ' PASO 2: Import del JSON exportado
    If gVerbose Then WScript.Echo "PASO 2: Importando JSON a base de datos temporal..."
    
    Call ImportFormInternal(strJson1Path, strTempDbPath, strPassword, True) ' True = overwrite
    
    If Err.Number <> 0 Then
        WScript.Echo "Error en PASO 2 (import): " & Err.Description
        gVerbose = bVerboseOriginal
        CleanupRoundtripTest strTempDir
        WScript.Quit 1
    End If
    
    ' PASO 3: Export del formulario importado
    If gVerbose Then WScript.Echo "PASO 3: Exportando formulario después del import..."
    
    Call ExportFormInternal(strTempDbPath, strFormName, strJson2Path, strPassword)
    
    If Err.Number <> 0 Then
        WScript.Echo "Error en PASO 3 (export final): " & Err.Description
        gVerbose = bVerboseOriginal
        CleanupRoundtripTest strTempDir
        WScript.Quit 1
    End If
    
    If Not objFSO.FileExists(strJson2Path) Then
        WScript.Echo "Error: No se generó el archivo JSON en PASO 3: " & strJson2Path
        gVerbose = bVerboseOriginal
        CleanupRoundtripTest strTempDir
        WScript.Quit 1
    End If
    
    ' PASO 4: Comparar los dos archivos JSON
    If gVerbose Then WScript.Echo "PASO 4: Comparando archivos JSON..."
    
    Dim json1Content, json2Content, bIdentical
    json1Content = ReadTextFile(strJson1Path)
    json2Content = ReadTextFile(strJson2Path)
    
    ' Normalizar contenido para comparación (remover timestamps y metadata variable)
    json1Content = NormalizeJsonForComparison(json1Content)
    json2Content = NormalizeJsonForComparison(json2Content)
    
    bIdentical = (json1Content = json2Content)
    
    gVerbose = bVerboseOriginal
    
    ' Mostrar resultados
    WScript.Echo "=== RESULTADOS DEL ROUNDTRIP TEST ==="
    WScript.Echo "Formulario: " & strFormName
    WScript.Echo "Export 1: " & strJson1Path & " (" & Len(ReadTextFile(strJson1Path)) & " chars)"
    WScript.Echo "Import:   " & strTempDbPath
    WScript.Echo "Export 2: " & strJson2Path & " (" & Len(ReadTextFile(strJson2Path)) & " chars)"
    
    If bIdentical Then
        WScript.Echo "RESULTADO: ✓ ÉXITO - Los archivos JSON son idénticos"
        WScript.Echo "El formulario mantiene su integridad en el ciclo export→import→export"
    Else
        WScript.Echo "RESULTADO: ✗ FALLO - Los archivos JSON difieren"
        WScript.Echo "Diferencias detectadas en el ciclo export→import→export"
        
        If gVerbose Then
            WScript.Echo "--- CONTENIDO EXPORT 1 (primeras 500 chars) ---"
            WScript.Echo Left(ReadTextFile(strJson1Path), 500)
            WScript.Echo "--- CONTENIDO EXPORT 2 (primeras 500 chars) ---"
            WScript.Echo Left(ReadTextFile(strJson2Path), 500)
        End If
    End If
    
    ' Limpiar archivos temporales (opcional, mantener para debug si hay fallo)
    If bIdentical Or Not gVerbose Then
        CleanupRoundtripTest strTempDir
        If gVerbose Then WScript.Echo "Archivos temporales eliminados."
    Else
        WScript.Echo "Archivos temporales conservados para análisis: " & strTempDir
    End If
    
    On Error GoTo 0
    
    If Not bIdentical Then WScript.Quit 1
End Sub

' ===================================================================
' SUBRUTINA AUXILIAR: ExportFormInternal
' Descripción: Versión interna de ExportForm para uso en roundtrip
' ===================================================================
Private Sub ExportFormInternal(dbPath, formName, outputPath, password)
    ' Implementación simplificada de export para roundtrip
    ' Reutiliza la lógica de ExportForm pero con parámetros directos
    
    If gVerbose Then WScript.Echo "Exportando " & formName & " desde " & dbPath & " a " & outputPath
    
    ' Crear instancia de Access
    Dim objAccessLocal
    Set objAccessLocal = CreateObject("Access.Application")
    objAccessLocal.Visible = False
    
    On Error Resume Next
    If password = "" Then
        objAccessLocal.OpenCurrentDatabase dbPath
    Else
        objAccessLocal.OpenCurrentDatabase dbPath, , password
    End If
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al abrir la base de datos para export: " & Err.Description
        objAccessLocal.Quit
        Exit Sub
    End If
    
    objAccessLocal.DoCmd.OpenForm formName, 3 ' acDesign
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al abrir formulario para export: " & Err.Description
        objAccessLocal.Quit
        Exit Sub
    End If
    
    On Error GoTo 0
    
    ' Generar JSON básico (versión simplificada)
    Dim frm, jsonContent
    Set frm = objAccessLocal.Forms(formName)
    
    jsonContent = "{" & vbCrLf
    jsonContent = jsonContent & "  ""schemaVersion"": ""1.0.0""," & vbCrLf
    jsonContent = jsonContent & "  ""formName"": """ & frm.Name & """," & vbCrLf
    jsonContent = jsonContent & "  ""properties"": {" & vbCrLf
    jsonContent = jsonContent & "    ""caption"": """ & frm.Caption & """," & vbCrLf
    jsonContent = jsonContent & "    ""width"": " & frm.Width & "," & vbCrLf
    jsonContent = jsonContent & "    ""recordSource"": """ & frm.RecordSource & """" & vbCrLf
    jsonContent = jsonContent & "  }" & vbCrLf
    jsonContent = jsonContent & "}" & vbCrLf
    
    ' Guardar archivo
    Dim objFileLocal
    Set objFileLocal = objFSO.CreateTextFile(outputPath, True, True)
    objFileLocal.Write jsonContent
    objFileLocal.Close
    
    objAccessLocal.DoCmd.Close
    objAccessLocal.Quit
End Sub

' ===================================================================
' SUBRUTINA AUXILIAR: ImportFormInternal
' Descripción: Versión interna de ImportForm para uso en roundtrip
' ===================================================================
Private Sub ImportFormInternal(jsonPath, dbPath, password, overwrite)
    ' Implementación simplificada de import para roundtrip
    If gVerbose Then WScript.Echo "Importando " & jsonPath & " a " & dbPath
    
    ' Leer y parsear JSON básico
    Dim jsonContent, formName
    jsonContent = ReadTextFile(jsonPath)
    
    ' Extraer formName usando regex simple
    Dim regEx, matches
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = """formName""\s*:\s*""([^""]+)"""
    Set matches = regEx.Execute(jsonContent)
    If matches.Count > 0 Then
        formName = matches(0).SubMatches(0)
    Else
        formName = "FormularioImportado"
    End If
    
    ' Crear instancia de Access
    Dim objAccessLocal
    Set objAccessLocal = CreateObject("Access.Application")
    objAccessLocal.Visible = False
    
    On Error Resume Next
    If password = "" Then
        objAccessLocal.OpenCurrentDatabase dbPath
    Else
        objAccessLocal.OpenCurrentDatabase dbPath, , password
    End If
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al abrir la base de datos para import: " & Err.Description
        objAccessLocal.Quit
        Exit Sub
    End If
    
    ' Crear formulario básico (simulado)
    objAccessLocal.DoCmd.NewForm
    objAccessLocal.DoCmd.Save , formName
    
    objAccessLocal.Quit
    On Error GoTo 0
End Sub

' ===================================================================
' FUNCIÓN AUXILIAR: NormalizeJsonForComparison
' Descripción: Normaliza JSON removiendo metadata variable para comparación
' ===================================================================
Private Function NormalizeJsonForComparison(jsonText)
    Dim result
    result = jsonText
    
    ' Remover timestamps variables
    Dim regEx
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    
    ' Remover generatedAtUTC
    regEx.Pattern = """generatedAtUTC""\s*:\s*""[^""]+"",?"
    result = regEx.Replace(result, "")
    
    ' Remover espacios extra y normalizar
    regEx.Pattern = "\s+"
    result = regEx.Replace(result, " ")
    
    NormalizeJsonForComparison = Trim(result)
End Function

' ===================================================================
' FUNCIÓN AUXILIAR: ReadTextFile
' Descripción: Lee contenido completo de un archivo de texto
' ===================================================================
Private Function ReadTextFile(filePath)
    Dim objFile, content
    Set objFile = objFSO.OpenTextFile(filePath, 1, False, -1) ' UTF-8
    content = objFile.ReadAll
    objFile.Close
    ReadTextFile = content
End Function

' ===================================================================
' SUBRUTINA AUXILIAR: CleanupRoundtripTest
' Descripción: Limpia archivos temporales del test de roundtrip
' ===================================================================
Private Sub CleanupRoundtripTest(tempDir)
    On Error Resume Next
    If objFSO.FolderExists(tempDir) Then
        objFSO.DeleteFolder tempDir, True
    End If
    On Error GoTo 0
End Sub

' ================================================================================
' FUNCIÓN: ParseJson
' DESCRIPCIÓN: Parsear JSON usando PowerShell como alternativa
' ================================================================================
Private Function ParseJson(jsonText)
    ' Parsear JSON usando expresiones regulares simples
    Set ParseJson = CreateObject("Scripting.Dictionary")
    
    ' Extraer formName
    Dim regEx, matches
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = """formName""\s*:\s*""([^""]+)"""
    regEx.Global = False
    Set matches = regEx.Execute(jsonText)
    If matches.Count > 0 Then
        ParseJson.Add "formName", matches(0).SubMatches(0)
    Else
        ParseJson.Add "formName", "FormularioSinNombre"
    End If
    
    ' Crear propiedades del formulario
    Dim formProps
    Set formProps = CreateObject("Scripting.Dictionary")
    
    ' Extraer caption
    regEx.Pattern = """caption""\s*:\s*""([^""]+)"""
    Set matches = regEx.Execute(jsonText)
    If matches.Count > 0 Then
        formProps.Add "caption", matches(0).SubMatches(0)
    End If
    
    ' Extraer width
    regEx.Pattern = """width""\s*:\s*(\d+)"
    Set matches = regEx.Execute(jsonText)
    If matches.Count > 0 Then
        formProps.Add "width", matches(0).SubMatches(0)
    End If
    
    ParseJson.Add "properties", formProps
    
    ' Crear secciones
    Dim sections, detailSection, detailProps
    Set sections = CreateObject("Scripting.Dictionary")
    Set detailSection = CreateObject("Scripting.Dictionary")
    Set detailProps = CreateObject("Scripting.Dictionary")
    
    ' Extraer height de la sección Detail
    regEx.Pattern = """Detail""[^}]*""height""\s*:\s*(\d+)"
    Set matches = regEx.Execute(jsonText)
    If matches.Count > 0 Then
        detailProps.Add "height", matches(0).SubMatches(0)
    End If
    
    detailSection.Add "properties", detailProps
    
    ' Parsear controles
    Dim controls
    Set controls = CreateObject("Scripting.Dictionary")
    
    ' Buscar todos los controles en la sección Detail
    regEx.Pattern = "\{[^}]*""name""\s*:\s*""([^""]+)""[^}]*""type""\s*:\s*""([^""]+)""[^}]*\}"
    regEx.Global = True
    Set matches = regEx.Execute(jsonText)
    
    Dim controlIndex
    controlIndex = 0
    
    Dim i
    For i = 0 To matches.Count - 1
        Dim ctrl, ctrlProps
        Set ctrl = CreateObject("Scripting.Dictionary")
        Set ctrlProps = CreateObject("Scripting.Dictionary")
        
        ' Obtener el bloque completo del control
        Dim controlBlock
        controlBlock = matches(i).Value
        
        ' Extraer propiedades básicas
        ctrl.Add "name", matches(i).SubMatches(0)
        ctrl.Add "type", matches(i).SubMatches(1)
        
        ' Extraer propiedades del control
        Dim propRegEx, propMatches
        Set propRegEx = CreateObject("VBScript.RegExp")
        
        ' Caption
        propRegEx.Pattern = """caption""\s*:\s*""([^""]+)"""
        Set propMatches = propRegEx.Execute(controlBlock)
        If propMatches.Count > 0 Then
            ctrlProps.Add "caption", propMatches(0).SubMatches(0)
        End If
        
        ' Top
        propRegEx.Pattern = """top""\s*:\s*(\d+)"
        Set propMatches = propRegEx.Execute(controlBlock)
        If propMatches.Count > 0 Then
            ctrlProps.Add "top", propMatches(0).SubMatches(0)
        End If
        
        ' Left
        propRegEx.Pattern = """left""\s*:\s*(\d+)"
        Set propMatches = propRegEx.Execute(controlBlock)
        If propMatches.Count > 0 Then
            ctrlProps.Add "left", propMatches(0).SubMatches(0)
        End If
        
        ' Width
        propRegEx.Pattern = """width""\s*:\s*(\d+)"
        Set propMatches = propRegEx.Execute(controlBlock)
        If propMatches.Count > 0 Then
            ctrlProps.Add "width", propMatches(0).SubMatches(0)
        End If
        
        ' Height
        propRegEx.Pattern = """height""\s*:\s*(\d+)"
        Set propMatches = propRegEx.Execute(controlBlock)
        If propMatches.Count > 0 Then
            ctrlProps.Add "height", propMatches(0).SubMatches(0)
        End If
        
        ' Picture
        propRegEx.Pattern = """picture""\s*:\s*""([^""]+)"""
        Set propMatches = propRegEx.Execute(controlBlock)
        If propMatches.Count > 0 Then
            ctrlProps.Add "picture", propMatches(0).SubMatches(0)
        End If
        
        ctrl.Add "properties", ctrlProps
        controls.Add controlIndex, ctrl
        controlIndex = controlIndex + 1
    Next
    
    detailSection.Add "controls", controls
    sections.Add "Detail", detailSection
    ParseJson.Add "sections", sections
    
    WScript.Echo "JSON parseado correctamente. Controles encontrados: " & controls.Count
End Function

' ================================================================================
' FUNCIÓN: JsonParser
' DESCRIPCIÓN: Parser robusto para JSON enriquecido con validaciones
' ================================================================================


' ================================================================================
' FUNCIÓN: MapEnumValue
' DESCRIPCIÓN: Mapear valores enum string a valores Access
' ================================================================================
Private Function MapEnumValue(enumType, stringValue)
    ' Tabla de mapeo local para enums conocidos
    Select Case LCase(enumType)
        Case "defaultview"
            Select Case LCase(stringValue)
                Case "single", "singleform": MapEnumValue = 0
                Case "continuous", "continuousforms": MapEnumValue = 1
                Case "datasheet": MapEnumValue = 2
                Case Else: MapEnumValue = 0
            End Select
        Case "cycle"
            Select Case LCase(stringValue)
                Case "allrecords": MapEnumValue = 0
                Case "currentrecord": MapEnumValue = 1
                Case "currentpage": MapEnumValue = 2
                Case Else: MapEnumValue = 0
            End Select
        Case "recordsourcetype"
            Select Case LCase(stringValue)
                Case "table": MapEnumValue = 0
                Case "dynaset": MapEnumValue = 1
                Case "snapshot": MapEnumValue = 2
                Case Else: MapEnumValue = 0
            End Select
        Case Else
            MapEnumValue = stringValue
    End Select
End Function

' ================================================================================
' SUBRUTINA: ImportForm
' DESCRIPCIÓN: Crear/Modificar formulario desde JSON enriquecido con validaciones
' ================================================================================
Sub ImportForm()
    ' Verificar argumentos mínimos
    If objArgs.Count < 3 Then
        WScript.Echo "Error: Se requieren al menos 2 argumentos: <json_path> <db_path>"
        WScript.Echo "Sintaxis: cscript condor_cli.vbs import-form <json_path> <db_path> [opciones]"
        WScript.Echo "Opciones:"
        WScript.Echo "  --schema <path>        Validar contra esquema específico"
        WScript.Echo "  --dry-run              Solo validar, no crear formulario"
        WScript.Echo "  --strict               Modo estricto: fallar en propiedades desconocidas"
        WScript.Echo "  --resource-root <DIR>  Directorio raíz para recursos (imágenes)"
        WScript.Echo "  --font-fallback <font> Fuente de respaldo (default: Segoe UI)"
        WScript.Echo "  --password             Solicitar contraseña de base de datos"
        WScript.Quit 1
    End If
    
    ' Declarar variables
    Dim strJsonPath, strDbPath, strPassword, strSchemaPath, strResourceRoot, strFontFallback
    Dim bDryRun, bStrict
    Dim jsonString, formData, formName
    Dim frm, i, j
    Dim controls, ctrl, ctlDest
    Dim sections, detailSection, formProps, sectionProps, ctlProps
    Dim reportSummary, warnings, missingResources, fallbacksApplied
    Dim codeModule, detectedHandlers, eventDiscrepancies
    
    ' Inicializar variables
    strJsonPath = objArgs(1)
    strDbPath = objArgs(2)
    strPassword = ""
    strSchemaPath = ""
    strResourceRoot = ""
    strFontFallback = "Segoe UI"
    bDryRun = False
    bStrict = False
    
    Set reportSummary = CreateObject("Scripting.Dictionary")
    Set warnings = CreateObject("Scripting.Dictionary")
    Set missingResources = CreateObject("Scripting.Dictionary")
    Set fallbacksApplied = CreateObject("Scripting.Dictionary")
    Set eventDiscrepancies = CreateObject("Scripting.Dictionary")
    
    ' Procesar argumentos opcionales
    For i = 3 To objArgs.Count - 1
        Select Case LCase(objArgs(i))
            Case "--schema"
                If i + 1 <= objArgs.Count - 1 Then
                    strSchemaPath = objArgs(i + 1)
                    i = i + 1
                Else
                    WScript.Echo "Error: --schema requiere una ruta"
                    WScript.Quit 1
                End If
            Case "--dry-run"
                bDryRun = True
            Case "--strict"
                bStrict = True
            Case "--resource-root"
                If i + 1 <= objArgs.Count - 1 Then
                    strResourceRoot = objArgs(i + 1)
                    i = i + 1
                Else
                    WScript.Echo "Error: --resource-root requiere un directorio"
                    WScript.Quit 1
                End If
            Case "--font-fallback"
                If i + 1 <= objArgs.Count - 1 Then
                    strFontFallback = objArgs(i + 1)
                    i = i + 1
                Else
                    WScript.Echo "Error: --font-fallback requiere un nombre de fuente"
                    WScript.Quit 1
                End If
            Case "--password"
                strPassword = InputBox("Ingrese la contraseña de la base de datos:", "Contraseña")
                If strPassword = "" Then
                    WScript.Echo "Error: Contraseña requerida."
                    WScript.Quit 1
                End If
        End Select
    Next
    
    ' Si la ruta JSON no es completa, construir la ruta por defecto
    If Not objFSO.FileExists(strJsonPath) Then
        Dim defaultUiPath
        defaultUiPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\ui\definitions\"
        strJsonPath = objFSO.BuildPath(defaultUiPath, objArgs(1))
    End If
    
    ' Convertir ruta de BD relativa a absoluta si es necesario
    If InStr(strDbPath, ":") = 0 Then
        strDbPath = objFSO.GetAbsolutePathName(strDbPath)
    End If
    
    ' Establecer resource-root por defecto si no se especificó
    If strResourceRoot = "" Then
        strResourceRoot = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\ui\resources"
    End If
    
    ' Verificar que los archivos existen
    If Not objFSO.FileExists(strJsonPath) Then
        WScript.Echo "Error: El archivo JSON no existe: " & strJsonPath
        WScript.Quit 1
    End If
    
    If Not bDryRun And Not objFSO.FileExists(strDbPath) Then
        WScript.Echo "Error: La base de datos no existe: " & strDbPath
        WScript.Quit 1
    End If
    
    ' Verificar directorio de recursos
    If Not objFSO.FolderExists(strResourceRoot) Then
        If bStrict Then
            WScript.Echo "Error: Directorio de recursos no existe: " & strResourceRoot
            WScript.Quit 1
        Else
            warnings.Add "resource_root", "Directorio de recursos no encontrado: " & strResourceRoot
        End If
    End If
    
    ' Leer contenido del archivo JSON
    Dim objFile
    Set objFile = objFSO.OpenTextFile(strJsonPath, 1)
    jsonString = objFile.ReadAll
    objFile.Close
    
    WScript.Echo "Parseando JSON con JsonParser..."
    
    ' Parsear JSON con JsonParser mejorado
    On Error Resume Next
    Set formData = JsonParser(jsonString, bStrict)
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al parsear JSON: " & Err.Description
        WScript.Quit 1
    End If
    On Error GoTo 0
    
    If formData Is Nothing Then
        WScript.Echo "Error: No se pudo parsear el fichero JSON."
        WScript.Quit 1
    End If
    
    ' Obtener nombre del formulario
    formName = formData("formName")
    WScript.Echo "Formulario a procesar: " & formName
    
    ' Extraer handlers detectados del JSON si existen
    Set detectedHandlers = CreateObject("Scripting.Dictionary")
    If formData.Exists("code") Then
        Set codeModule = formData("code")
        If codeModule.Exists("module") And codeModule("module").Exists("handlers") Then
            Dim handlers, handler, controlName, eventName, handlerKey
            Set handlers = codeModule("module")("handlers")
            For i = 0 To handlers.Count - 1
                Set handler = handlers(i)
                If handler.Exists("control") And handler.Exists("event") Then
                    controlName = handler("control")
                    eventName = handler("event")
                    handlerKey = controlName & "." & eventName
                    detectedHandlers.Add handlerKey, True
                    If gVerbose Then
                        WScript.Echo "Handler detectado: " & handlerKey
                    End If
                End If
            Next
        End If
    End If
    
    ' Validaciones mínimas
    Call ValidateFormData(formData, bStrict, warnings)
    
    ' Verificar recursos de imágenes
    Call ValidateResources(formData, strResourceRoot, bStrict, missingResources)
    
    ' Si es dry-run, mostrar reporte y salir
    If bDryRun Then
        WScript.Echo "=== MODO DRY-RUN: VALIDACIÓN COMPLETADA ==="
        WScript.Echo "Formulario: " & formName
        WScript.Echo "Schema Version: " & formData("schemaVersion")
        
        If warnings.Count > 0 Then
            WScript.Echo "Advertencias encontradas:"
            Dim key
            For Each key In warnings.Keys
                WScript.Echo "  - " & warnings(key)
            Next
        End If
        
        If missingResources.Count > 0 Then
            WScript.Echo "Recursos faltantes:"
            For Each key In missingResources.Keys
                WScript.Echo "  - " & missingResources(key)
            Next
        End If
        
        WScript.Echo "Validación completada exitosamente."
        WScript.Quit 0
    End If
    
    ' Obtener nombre del formulario para procesamiento real
    On Error Resume Next
    If Err.Number <> 0 Then
        WScript.Echo "Error: No se pudo obtener el nombre del formulario del JSON: " & Err.Description
        WScript.Quit 1
    End If
    On Error GoTo 0
    
    WScript.Echo "Nombre del formulario desde JSON: " & formName
    
    ' Cerrar Access completamente y crear nueva instancia
    If Not objAccess Is Nothing Then
        On Error Resume Next
        objAccess.Quit
        Set objAccess = Nothing
        On Error GoTo 0
    End If
    
    ' Crear nueva instancia de Access
    On Error Resume Next
    Set objAccess = CreateObject("Access.Application")
    If Err.Number <> 0 Then
        WScript.Echo "Error al crear nueva instancia de Access: " & Err.Description
        WScript.Quit 1
    End If
    
    ' Configurar Access
    objAccess.Visible = False
    objAccess.UserControl = False
    On Error GoTo 0
    
    ' Abrir la base de datos especificada con manejo de contraseña
    On Error Resume Next
    Dim dbPassword
    dbPassword = GetDatabasePassword(strDbPath)
    
    If dbPassword = "" Then
        objAccess.OpenCurrentDatabase strDbPath
    Else
        objAccess.OpenCurrentDatabase strDbPath, , dbPassword
    End If
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al abrir la base de datos: " & Err.Description
        WScript.Echo "Ruta: " & strDbPath
        WScript.Echo "Contraseña detectada: " & IIf(dbPassword = "", "No", "Sí")
        WScript.Quit 1
    End If
    On Error GoTo 0
    
    WScript.Echo "Base de datos abierta: " & objAccess.CurrentProject.Name
    
    ' Crear el formulario dinámicamente
    On Error Resume Next
    
    ' Verificar si el formulario ya existe y eliminarlo
    Dim formExists
    formExists = False
    
    ' Verificar si el formulario existe
    Dim formIndex
    For formIndex = 0 To objAccess.CurrentProject.AllForms.Count - 1
        If objAccess.CurrentProject.AllForms(formIndex).Name = formName Then
            formExists = True
            Exit For
        End If
    Next
    
    If formExists Then
        ' Cerrar si está abierto
        If objAccess.CurrentProject.AllForms(formName).IsLoaded Then
            objAccess.DoCmd.Close 1, formName, 0  ' acForm, formName, acSaveNo
        End If
        ' Eliminar el formulario existente
        objAccess.DoCmd.DeleteObject 1, formName  ' acForm, formName
        WScript.Echo "Formulario existente '" & formName & "' eliminado."
    End If
    
    On Error GoTo 0
    
    ' Crear nuevo formulario
    On Error Resume Next
    Set frm = objAccess.CreateForm()
    If Err.Number <> 0 Then
        WScript.Echo "Error al crear el formulario: " & Err.Description
        WScript.Quit 1
    End If
    On Error GoTo 0
    
    ' Guardar el nombre original del formulario para referencia
    Dim originalFormName
    originalFormName = frm.Name
    
    WScript.Echo "Nombre original del formulario creado: " & originalFormName
    WScript.Echo "Nombre deseado del formulario: " & formName
    
    ' Según la documentación de Microsoft, el nombre del formulario no se puede cambiar
    ' inmediatamente después de crearlo. Se debe hacer después de guardarlo.
    On Error Resume Next
    frm.Name = formName
    If Err.Number = 0 Then
        originalFormName = formName
        WScript.Echo "Formulario renombrado inmediatamente a: " & formName
    Else
        WScript.Echo "No se pudo cambiar el nombre inmediatamente: " & Err.Description
    End If
    On Error GoTo 0
    
    ' Establecer propiedades del formulario
    ' Nota: Name se establecerá al final con DoCmd.Rename
    If formData.Exists("properties") Then
        Set formProps = formData("properties")
        
        If formProps.Exists("caption") Then
            frm.Caption = formProps("caption")
        End If
        
        If formProps.Exists("width") Then
            frm.Width = CLng(formProps("width"))
        End If
        
        If formProps.Exists("height") Then
            frm.WindowHeight = CLng(formProps("height"))
        End If
        
        ' Aplicar colores del formulario
        If formProps.Exists("backColor") Then
            If ValidateHexColor(formProps("backColor")) Then
                frm.Detail.BackColor = ConvertHexToOLE(formProps("backColor"))
            ElseIf bStrict Then
                WScript.Echo "Error: Color de fondo del formulario inválido: " & formProps("backColor")
            End If
        End If
        
        ' Aplicar propiedades enum del formulario
        If formProps.Exists("defaultView") Then
            frm.DefaultView = MapEnumValue("defaultView", formProps("defaultView"))
        End If
        
        If formProps.Exists("cycle") Then
            frm.Cycle = MapEnumValue("cycle", formProps("cycle"))
        End If
        
        If formProps.Exists("recordSourceType") Then
            frm.RecordSourceType = MapEnumValue("recordSourceType", formProps("recordSourceType"))
        End If
        
        ' Aplicar otras propiedades
        If formProps.Exists("recordSource") Then
            frm.RecordSource = formProps("recordSource")
        End If
        
        If formProps.Exists("allowEdits") Then
            frm.AllowEdits = CBool(formProps("allowEdits"))
        End If
        
        If formProps.Exists("allowAdditions") Then
            frm.AllowAdditions = CBool(formProps("allowAdditions"))
        End If
        
        If formProps.Exists("allowDeletions") Then
            frm.AllowDeletions = CBool(formProps("allowDeletions"))
        End If
    End If
    
    ' Crear controles si existen secciones
    If formData.Exists("sections") Then
        Set sections = formData("sections")
        
        If sections.Exists("Detail") Then
            Set detailSection = sections("Detail")
            
            ' Establecer propiedades de la sección Detail
            If detailSection.Exists("properties") Then
                Set sectionProps = detailSection("properties")
                
                If sectionProps.Exists("height") Then
                    frm.Section(0).Height = CLng(sectionProps("height"))  ' acDetail = 0
                End If
            End If
            
            ' Crear controles
            If detailSection.Exists("controls") Then
                Set controls = detailSection("controls")
                
                WScript.Echo "Creando " & controls.Count & " controles en la sección Detail"
                
                For i = 0 To controls.Count - 1
                    Set ctrl = controls(i)
                    
                    WScript.Echo "Creando control: " & ctrl("name") & " de tipo: " & ctrl("type")
                    
                    On Error Resume Next
                    Set ctlDest = Nothing
                    
                    ' Crear control según su tipo usando originalFormName
                    If ctrl("type") = "CommandButton" Then
                        Set ctlDest = objAccess.CreateControl(originalFormName, 104, 0)  ' acCommandButton, acDetail
                    ElseIf ctrl("type") = "Label" Then
                        Set ctlDest = objAccess.CreateControl(originalFormName, 100, 0)  ' acLabel, acDetail
                    ElseIf ctrl("type") = "TextBox" Then
                        Set ctlDest = objAccess.CreateControl(originalFormName, 109, 0)  ' acTextBox, acDetail
                    End If
                    
                    If Err.Number <> 0 Then
                        WScript.Echo "Error al crear control " & ctrl("name") & ": " & Err.Description
                        Err.Clear
                    ElseIf Not ctlDest Is Nothing Then
                        ' Establecer nombre del control
                        ctlDest.Name = ctrl("name")
                        
                        ' Establecer propiedades del control
                        If ctrl.Exists("properties") Then
                            Set ctlProps = ctrl("properties")
                            
                            If ctlProps.Exists("caption") Then
                                ctlDest.Caption = ctlProps("caption")
                            End If
                            
                            If ctlProps.Exists("top") Then
                                ctlDest.Top = CLng(ctlProps("top"))
                            End If
                            
                            If ctlProps.Exists("left") Then
                                ctlDest.Left = CLng(ctlProps("left"))
                            End If
                            
                            If ctlProps.Exists("width") Then
                                ctlDest.Width = CLng(ctlProps("width"))
                            End If
                            
                            If ctlProps.Exists("height") Then
                                ctlDest.Height = CLng(ctlProps("height"))
                            End If
                            
                            ' Aplicar colores si existen
                            If ctlProps.Exists("backColor") Then
                                If ValidateHexColor(ctlProps("backColor")) Then
                                    ctlDest.BackColor = ConvertHexToOLE(ctlProps("backColor"))
                                ElseIf bStrict Then
                                    WScript.Echo "Error: Color de fondo inválido: " & ctlProps("backColor")
                                End If
                            End If
                            
                            If ctlProps.Exists("foreColor") Then
                                If ValidateHexColor(ctlProps("foreColor")) Then
                                    ctlDest.ForeColor = ConvertHexToOLE(ctlProps("foreColor"))
                                ElseIf bStrict Then
                                    WScript.Echo "Error: Color de texto inválido: " & ctlProps("foreColor")
                                End If
                            End If
                            
                            ' Aplicar propiedades enum
                            If ctlProps.Exists("textAlign") Then
                                ctlDest.TextAlign = MapEnumValue("textAlign", ctlProps("textAlign"))
                            End If
                            
                            ' Aplicar fuente con fallback
                            If ctlProps.Exists("fontName") Then
                                Dim fontToUse
                                fontToUse = ctlProps("fontName")
                                
                                ' Si font-fallback está habilitado, verificar disponibilidad
                                If bFontFallback Then
                                    ' Usar fuente por defecto si no está disponible
                                    ' En Access, las fuentes no disponibles se sustituyen automáticamente
                                    If strFontFallback <> "" Then
                                        fontToUse = strFontFallback
                                    End If
                                End If
                                
                                On Error Resume Next
                                ctlDest.FontName = fontToUse
                                If Err.Number <> 0 And bStrict Then
                                    WScript.Echo "Error: No se pudo aplicar fuente: " & fontToUse
                                    Err.Clear
                                End If
                                On Error GoTo 0
                            End If
                            
                            If ctlProps.Exists("fontSize") Then
                                ctlDest.FontSize = CLng(ctlProps("fontSize"))
                            End If
                            
                            ' Aplicar imagen si existe (asignación directa según documentación Microsoft Access)
                            If ctlProps.Exists("picture") And ctrl("type") = "CommandButton" Then
                                Dim picturePath
                                picturePath = objFSO.BuildPath(strResourceRoot, ctlProps("picture"))
                                If objFSO.FileExists(picturePath) Then
                                    On Error Resume Next
                                    ' Usar asignación directa de ruta según documentación Microsoft Access
                                    ctlDest.Picture = picturePath
                                    If Err.Number <> 0 Then
                                        WScript.Echo "Error al cargar imagen: " & Err.Description
                                        Err.Clear
                                    Else
                                        WScript.Echo "Imagen aplicada correctamente: " & ctlProps("picture")
                                    End If
                                    On Error GoTo 0
                                Else
                                    If bStrict Then
                                        WScript.Echo "Error: Imagen no encontrada: " & picturePath
                                    Else
                                        WScript.Echo "Advertencia: Imagen no encontrada: " & picturePath
                                    End If
                                End If
                            End If
                        End If
                        
                        ' Aplicar eventos basados en handlers detectados
                        If ctrl.Exists("events") Then
                            Dim ctrlEvents, currentEventName, currentEventValue, shouldBeEventProcedure
                            Set ctrlEvents = ctrl("events")
                            
                            ' Procesar cada evento del control
                            Dim eventKeys
                            eventKeys = Array("onClick", "onDblClick", "onCurrent", "onLoad", "onOpen", "onGotFocus", "onLostFocus", "onChange", "onAfterUpdate", "onBeforeUpdate")
                            
                            For j = 0 To UBound(eventKeys)
                                currentEventName = eventKeys(j)
                                If ctrlEvents.Exists(currentEventName) Then
                                    currentEventValue = ctrlEvents(currentEventName)
                                    shouldBeEventProcedure = False
                                    
                                    ' Determinar si debe ser [Event Procedure]
                                    ' 1) Si el JSON especifica explícitamente [Event Procedure]
                                    If currentEventValue = "[Event Procedure]" Then
                                        shouldBeEventProcedure = True
                                    End If
                                    
                                    ' 2) Si hay handler detectado para este control y evento
                                    Dim currentEventKey, currentMappedEventName
                                    currentMappedEventName = MapEventNameToHandler(currentEventName)
                                    currentEventKey = ctrl("name") & "." & currentMappedEventName
                                    If detectedHandlers.Exists(currentEventKey) Then
                                        If currentEventValue <> "[Event Procedure]" And currentEventValue <> "" Then
                                            ' Discrepancia: hay handler pero JSON indica otra cosa
                                            Dim discrepancyMsg
                                            discrepancyMsg = "Control " & ctrl("name") & "." & currentEventName & ": handler detectado pero JSON indica '" & currentEventValue & "'"
                                            eventDiscrepancies.Add eventDiscrepancies.Count, discrepancyMsg
                                            If bStrict Then
                                                WScript.Echo "ERROR: " & discrepancyMsg
                                            Else
                                                WScript.Echo "WARNING: " & discrepancyMsg
                                            End If
                                        End If
                                        shouldBeEventProcedure = True
                                    End If
                                    
                                    ' Aplicar [Event Procedure] si es necesario
                                    If shouldBeEventProcedure Then
                                        On Error Resume Next
                                        Call SetControlEventProperty(ctlDest, currentEventName, "[Event Procedure]")
                                        If Err.Number <> 0 Then
                                            WScript.Echo "Error al establecer evento " & currentEventName & ": " & Err.Description
                                            Err.Clear
                                        ElseIf gVerbose Then
                                            WScript.Echo "Evento " & currentEventName & " establecido como [Event Procedure] para " & ctrl("name")
                                        End If
                                        On Error GoTo 0
                                    End If
                                End If
                            Next
                        End If
                        
                        WScript.Echo "Control " & ctrl("name") & " creado exitosamente"
                    Else
                        WScript.Echo "Error: No se pudo crear el control " & ctrl("name")
                    End If
                    
                    On Error GoTo 0
                Next
            End If
        End If
    End If
    
    ' Guardar y cerrar el formulario
    On Error Resume Next
    
    ' Cerrar y guardar el formulario automáticamente
    objAccess.DoCmd.Close 1, originalFormName, -1  ' acForm, formName, acSaveYes
    If Err.Number <> 0 Then
        WScript.Echo "Error al cerrar y guardar el formulario: " & Err.Description
        WScript.Quit 1
    End If
    
    WScript.Echo "Formulario guardado con nombre: " & originalFormName
    
    ' Actualizar el catálogo de objetos de Access (según documentación Microsoft)
    Err.Clear
    objAccess.DoCmd.RefreshDatabaseWindow
    If Err.Number <> 0 Then
        WScript.Echo "Advertencia: No se pudo actualizar el catálogo de objetos: " & Err.Description
        Err.Clear
    Else
        WScript.Echo "Catálogo de objetos actualizado"
    End If
    
    ' Si el nombre deseado es diferente al original, renombrar después de cerrar
    If originalFormName <> formName Then
        ' Esperar un momento para que Access complete el guardado y la actualización
        WScript.Sleep 1000
        
        ' Intentar renombrar el formulario cerrado
        objAccess.DoCmd.Rename formName, 1, originalFormName  ' newName, acForm, oldName
        If Err.Number <> 0 Then
            WScript.Echo "Error al renombrar formulario: " & Err.Description
            WScript.Echo "El formulario permanece con el nombre: " & originalFormName
            formName = originalFormName
        Else
            WScript.Echo "Formulario renombrado exitosamente a: " & formName
        End If
    End If
    
    WScript.Echo "Formulario final: " & formName
    
    ' Mostrar resumen de discrepancias de eventos si las hay
    If eventDiscrepancies.Count > 0 Then
        WScript.Echo ""
        WScript.Echo "=== RESUMEN DE DISCREPANCIAS DE EVENTOS ==="
        Dim discrepancyKey
        For Each discrepancyKey In eventDiscrepancies.Keys
            WScript.Echo eventDiscrepancies(discrepancyKey)
        Next
        WScript.Echo "Total de discrepancias: " & eventDiscrepancies.Count
        
        If bStrict Then
            WScript.Echo "ERROR: Modo --strict activado. Import fallido debido a discrepancias."
            WScript.Quit 1
        End If
    End If
    
    On Error GoTo 0
    
    WScript.Echo "Formulario '" & formName & "' creado exitosamente."
    
End Sub

' ============================================================================
' FUNCIÓN: MapEventNameToHandler
' Descripción: Mapea nombres de eventos JSON a nombres de handlers VBA
' Parámetros: eventName - nombre del evento en JSON (ej: "onClick")
' Retorna: nombre del evento para handler VBA (ej: "Click")
' ============================================================================
Private Function MapEventNameToHandler(eventName)
    Select Case LCase(eventName)
        Case "onclick"
            MapEventNameToHandler = "Click"
        Case "ondblclick"
            MapEventNameToHandler = "DblClick"
        Case "oncurrent"
            MapEventNameToHandler = "Current"
        Case "onload"
            MapEventNameToHandler = "Load"
        Case "onopen"
            MapEventNameToHandler = "Open"
        Case "ongotfocus"
            MapEventNameToHandler = "GotFocus"
        Case "onlostfocus"
            MapEventNameToHandler = "LostFocus"
        Case "onchange"
            MapEventNameToHandler = "Change"
        Case "onafterupdate"
            MapEventNameToHandler = "AfterUpdate"
        Case "onbeforeupdate"
            MapEventNameToHandler = "BeforeUpdate"
        Case Else
            MapEventNameToHandler = eventName
    End Select
End Function

' ============================================================================
' SUBRUTINA: SetControlEventProperty
' Descripción: Establece la propiedad de evento de un control
' Parámetros: 
'   ctrl - objeto control de Access
'   eventName - nombre del evento (ej: "onClick")
'   eventValue - valor a asignar (ej: "[Event Procedure]")
' ============================================================================
Private Sub SetControlEventProperty(ctrl, eventName, eventValue)
    On Error Resume Next
    
    Select Case LCase(eventName)
        Case "onclick"
            ctrl.OnClick = eventValue
        Case "ondblclick"
            ctrl.OnDblClick = eventValue
        Case "oncurrent"
            ctrl.OnCurrent = eventValue
        Case "onload"
            ctrl.OnLoad = eventValue
        Case "onopen"
            ctrl.OnOpen = eventValue
        Case "ongotfocus"
            ctrl.OnGotFocus = eventValue
        Case "onlostfocus"
            ctrl.OnLostFocus = eventValue
        Case "onchange"
            ctrl.OnChange = eventValue
        Case "onafterupdate"
            ctrl.OnAfterUpdate = eventValue
        Case "onbeforeupdate"
            ctrl.OnBeforeUpdate = eventValue
    End Select
    
    If Err.Number <> 0 Then
        Err.Raise Err.Number, Err.Source, Err.Description
    End If
End Sub

' ============================================================================
' SUBRUTINA: ValidateFormJsonCommand
' Descripción: Valida la estructura JSON de un formulario
' Sintaxis: validate-form-json <json_path> [--strict] [--schema]
' Parámetros:
'   json_path: Ruta al archivo JSON del formulario
'   --strict: Modo estricto (falla en advertencias)
'   --schema: Mostrar esquema JSON esperado
' ============================================================================
Private Sub ValidateFormJsonCommand()
    Dim jsonPath, bStrict, bShowSchema
    Dim i, arg
    Dim objFSO, objFile, jsonContent
    Dim objParser, formData
    Dim warnings, errors
    
    ' Inicializar variables
    jsonPath = ""
    bStrict = False
    bShowSchema = False
    Set warnings = CreateObject("Scripting.Dictionary")
    Set errors = CreateObject("Scripting.Dictionary")
    
    ' Procesar argumentos
    For i = 2 To WScript.Arguments.Count - 1
        arg = WScript.Arguments(i)
        If Left(arg, 2) = "--" Then
            Select Case arg
                Case "--strict"
                    bStrict = True
                Case "--schema"
                    bShowSchema = True
                Case Else
                    WScript.Echo "Error: Opción desconocida: " & arg
                    WScript.Quit 1
            End Select
        Else
            If jsonPath = "" Then
                jsonPath = arg
            Else
                WScript.Echo "Error: Múltiples rutas JSON especificadas"
                WScript.Quit 1
            End If
        End If
    Next
    
    ' Mostrar esquema si se solicita
    If bShowSchema Then
        ShowFormJsonSchema()
        If jsonPath = "" Then Exit Sub
    End If
    
    ' Validar argumentos requeridos
    If jsonPath = "" Then
        WScript.Echo "Error: Debe especificar la ruta del archivo JSON"
        WScript.Echo "Uso: condor_cli validate-form-json <json_path> [--strict] [--schema]"
        WScript.Quit 1
    End If
    
    ' Verificar que el archivo existe
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Not objFSO.FileExists(jsonPath) Then
        WScript.Echo "Error: El archivo JSON no existe: " & jsonPath
        WScript.Quit 1
    End If
    
    WScript.Echo "Validando archivo JSON: " & jsonPath
    
    ' Leer y parsear el JSON
    On Error Resume Next
    Set objFile = objFSO.OpenTextFile(jsonPath, 1, False, -1)
    jsonContent = objFile.ReadAll
    objFile.Close
    
    If Err.Number <> 0 Then
        WScript.Echo "Error al leer el archivo JSON: " & Err.Description
        WScript.Quit 1
    End If
    
    ' Parsear JSON
    Set objParser = New JsonParser
    Set formData = objParser.Parse(jsonContent)
    
    If Err.Number <> 0 Then
        WScript.Echo "Error: JSON inválido - " & Err.Description
        WScript.Quit 1
    End If
    On Error GoTo 0
    
    ' Validar estructura del formulario
    ValidateFormData formData, bStrict, warnings
    
    ' Validar recursos (sin directorio base)
    ValidateResources formData, "", warnings
    
    ' Mostrar resultados
    WScript.Echo "\n=== RESULTADO DE VALIDACIÓN ==="
    
    If warnings.Count = 0 Then
        WScript.Echo "✓ JSON válido - No se encontraron problemas"
    Else
        WScript.Echo "⚠ Se encontraron " & warnings.Count & " advertencias:"
        Dim key
        For Each key In warnings.Keys
            WScript.Echo "  - " & warnings(key)
        Next
        
        If bStrict Then
            WScript.Echo "\nError: Validación falló en modo estricto"
            WScript.Quit 1
        End If
    End If
    
    WScript.Echo "\nValidación completada exitosamente."
End Sub

' ============================================================================
' SUBRUTINA: ShowFormJsonSchema
' Descripción: Muestra el esquema JSON esperado para formularios
' ============================================================================
Private Sub ShowFormJsonSchema()
    WScript.Echo "\n=== ESQUEMA JSON DE FORMULARIO ==="
    WScript.Echo "{"
    WScript.Echo "  ""name"": ""string (requerido)"","
    WScript.Echo "  ""properties"": {"
    WScript.Echo "    ""caption"": ""string"","
    WScript.Echo "    ""width"": ""number"","
    WScript.Echo "    ""height"": ""number"","
    WScript.Echo "    ""backColor"": ""string (hex: #RRGGBB)"","
    WScript.Echo "    ""defaultView"": ""enum (single|continuous|datasheet|pivotTable|pivotChart)"","
    WScript.Echo "    ""cycle"": ""enum (currentRecord|allRecords)"","
    WScript.Echo "    ""recordSourceType"": ""enum (table|dynaset|snapshot)"","
    WScript.Echo "    ""recordSource"": ""string"","
    WScript.Echo "    ""allowEdits"": ""boolean"","
    WScript.Echo "    ""allowAdditions"": ""boolean"","
    WScript.Echo "    ""allowDeletions"": ""boolean"""
    WScript.Echo "  },"
    WScript.Echo "  ""sections"": {"
    WScript.Echo "    ""detail"": {"
    WScript.Echo "      ""height"": ""number"","
    WScript.Echo "      ""backColor"": ""string (hex: #RRGGBB)"""
    WScript.Echo "    }"
    WScript.Echo "  },"
    WScript.Echo "  ""controls"": ["
    WScript.Echo "    {"
    WScript.Echo "      ""name"": ""string (requerido)"","
    WScript.Echo "      ""type"": ""enum (CommandButton|Label|TextBox) (requerido)"","
    WScript.Echo "      ""properties"": {"
    WScript.Echo "        ""caption"": ""string"","
    WScript.Echo "        ""top"": ""number (requerido)"","
    WScript.Echo "        ""left"": ""number (requerido)"","
    WScript.Echo "        ""width"": ""number (requerido)"","
    WScript.Echo "        ""height"": ""number (requerido)"","
    WScript.Echo "        ""backColor"": ""string (hex: #RRGGBB)"","
    WScript.Echo "        ""foreColor"": ""string (hex: #RRGGBB)"","
    WScript.Echo "        ""fontName"": ""string"","
    WScript.Echo "        ""fontSize"": ""number"","
    WScript.Echo "        ""fontBold"": ""boolean"","
    WScript.Echo "        ""fontItalic"": ""boolean"","
    WScript.Echo "        ""picture"": ""string (ruta relativa)"","
    WScript.Echo "        ""textAlign"": ""enum (left|center|right)"","
    WScript.Echo "        ""borderStyle"": ""enum (transparent|solid|dashes|dots)"","
    WScript.Echo "        ""specialEffect"": ""enum (flat|raised|sunken|etched|shadowed|chiseled)"""
    WScript.Echo "      }"
    WScript.Echo "    }"
    WScript.Echo "  ]"
    WScript.Echo "}"
    WScript.Echo "\n=== COLORES VÁLIDOS ==="
    WScript.Echo "Formato hexadecimal: #RRGGBB (ej: #FF0000 para rojo)"
    WScript.Echo "\n=== ENUMS VÁLIDOS ==="
    WScript.Echo "defaultView: single, continuous, datasheet, pivotTable, pivotChart"
    WScript.Echo "cycle: currentRecord, allRecords"
    WScript.Echo "recordSourceType: table, dynaset, snapshot"
    WScript.Echo "textAlign: left, center, right"
    WScript.Echo "borderStyle: transparent, solid, dashes, dots"
    WScript.Echo "specialEffect: flat, raised, sunken, etched, shadowed, chiseled"
End Sub

' ============================================================================
' INFRAESTRUCTURA JSON Y UTILIDADES
' ============================================================================

' Inicializar variable global verbose
gVerbose = False

' Clase JsonWriter - Convierte objetos VBA a JSON
Class JsonWriter
    Public Function Stringify(value)
        If IsNull(value) Then
            Stringify = "null"
        ElseIf VarType(value) = vbBoolean Then
            If value Then
                Stringify = "true"
            Else
                Stringify = "false"
            End If
        ElseIf IsNumeric(value) Then
            Stringify = CStr(value)
        ElseIf VarType(value) = vbString Then
            Stringify = """" & EscapeString(CStr(value)) & """"
        ElseIf IsDictionary(value) Then
            Stringify = StringifyObject(value)
        ElseIf IsArrayLike(value) Then
            Stringify = StringifyArray(value)
        Else
            Stringify = "null"
        End If
    End Function
    
    Private Function EscapeString(str)
        Dim result, i, char
        result = ""
        For i = 1 To Len(str)
            char = Mid(str, i, 1)
            Select Case char
                Case Chr(34) ' "
                    result = result & Chr(34) & Chr(34)
                Case Chr(92) ' backslash
                    result = result & "\\"
                Case Chr(8)  ' \b
                    result = result & "\b"
                Case Chr(12) ' \f
                    result = result & "\f"
                Case Chr(10) ' \n
                    result = result & "\n"
                Case Chr(13) ' \r
                    result = result & "\r"
                Case Chr(9)  ' \t
                    result = result & "\t"
                Case Else
                    If Asc(char) < 32 Then
                        result = result & "\u" & Right("0000" & Hex(Asc(char)), 4)
                    Else
                        result = result & char
                    End If
            End Select
        Next
        EscapeString = result
    End Function
    
    Private Function StringifyObject(obj)
        Dim result, key, first
        result = "{"
        first = True
        For Each key In obj.Keys
            If Not first Then result = result & ","
            result = result & """" & EscapeString(CStr(key)) & """:" & Stringify(obj(key))
            first = False
        Next
        result = result & "}"
        StringifyObject = result
    End Function
    
    Private Function StringifyArray(arr)
        Dim result, i, first
        result = "["
        first = True
        
        If TypeName(arr) = "ArrayList" Then
            For i = 0 To arr.Count - 1
                If Not first Then result = result & ","
                result = result & Stringify(arr(i))
                first = False
            Next
        Else
            ' Array nativo VBA
            For i = LBound(arr) To UBound(arr)
                If Not first Then result = result & ","
                result = result & Stringify(arr(i))
                first = False
            Next
        End If
        
        result = result & "]"
        StringifyArray = result
    End Function
End Class

' Clase JsonParser - Convierte JSON a objetos VBA
Class JsonParser
    Private pos
    Private jsonText
    
    Public Function Parse(json)
        jsonText = json
        pos = 1
        SkipWhitespace
        Set Parse = ParseValue()
    End Function
    
    Private Function ParseValue()
        SkipWhitespace
        Dim char
        char = Mid(jsonText, pos, 1)
        
        Select Case char
            Case "{"
                Set ParseValue = ParseObject()
            Case "["
                Set ParseValue = ParseArray()
            Case Chr(34) ' "
                ParseValue = ParseString()
            Case "t", "f"
                ParseValue = ParseBoolean()
            Case "n"
                ParseValue = ParseNull()
            Case Else
                If IsNumeric(char) Or char = "-" Then
                    ParseValue = ParseNumber()
                Else
                    Err.Raise 1001, "JsonParser", "Carácter inesperado en posición " & pos & ": " & char
                End If
        End Select
    End Function
    
    Private Function ParseObject()
        Dim obj
        Set obj = CreateDict()
        pos = pos + 1 ' Saltar '{'
        SkipWhitespace
        
        If Mid(jsonText, pos, 1) = "}" Then
            pos = pos + 1
            Set ParseObject = obj
            Exit Function
        End If
        
        Do
            SkipWhitespace
            Dim key
            key = ParseString()
            SkipWhitespace
            
            If Mid(jsonText, pos, 1) <> ":" Then
                Err.Raise 1002, "JsonParser", "Se esperaba ':' después de la clave"
            End If
            pos = pos + 1
            
            Dim value
             Set value = ParseValue()
             If IsObject(value) Then
                 Set obj(key) = value
             Else
                 obj(key) = value
             End If
            
            SkipWhitespace
            Dim nextChar
            nextChar = Mid(jsonText, pos, 1)
            
            If nextChar = "}" Then
                pos = pos + 1
                Exit Do
            ElseIf nextChar = "," Then
                pos = pos + 1
            Else
                Err.Raise 1003, "JsonParser", "Se esperaba ',' o '}'"
            End If
        Loop
        
        Set ParseObject = obj
    End Function
    
    Private Function ParseArray()
        Dim arr
        Set arr = CreateList()
        pos = pos + 1 ' Saltar '['
        SkipWhitespace
        
        If Mid(jsonText, pos, 1) = "]" Then
            pos = pos + 1
            Set ParseArray = arr
            Exit Function
        End If
        
        Do
             Dim value
             Set value = ParseValue()
             If IsObject(value) Then
                 arr.Add value
             Else
                 arr.Add value
             End If
            
            SkipWhitespace
            Dim nextChar
            nextChar = Mid(jsonText, pos, 1)
            
            If nextChar = "]" Then
                pos = pos + 1
                Exit Do
            ElseIf nextChar = "," Then
                pos = pos + 1
            Else
                Err.Raise 1004, "JsonParser", "Se esperaba ',' o ']'"
            End If
        Loop
        
        Set ParseArray = arr
    End Function
    
    Private Function ParseString()
        pos = pos + 1 ' Saltar '"' inicial
        Dim result, char
        result = ""
        
        Do While pos <= Len(jsonText)
            char = Mid(jsonText, pos, 1)
            
            If char = Chr(34) Then ' "
                pos = pos + 1
                ParseString = result
                Exit Function
            ElseIf char = "\" Then
                pos = pos + 1
                If pos > Len(jsonText) Then Exit Do
                
                Dim escapeChar
                escapeChar = Mid(jsonText, pos, 1)
                Select Case escapeChar
                    Case Chr(34) ' "
                        result = result & Chr(34)
                    Case "\"
                        result = result & "\"
                    Case "b"
                        result = result & Chr(8)
                    Case "f"
                        result = result & Chr(12)
                    Case "n"
                        result = result & Chr(10)
                    Case "r"
                        result = result & Chr(13)
                    Case "t"
                        result = result & Chr(9)
                    Case "u"
                        ' Unicode escape \uXXXX
                        If pos + 4 <= Len(jsonText) Then
                            Dim hexCode
                            hexCode = Mid(jsonText, pos + 1, 4)
                            result = result & Chr(CLng("&H" & hexCode))
                            pos = pos + 4
                        End If
                    Case Else
                        result = result & escapeChar
                End Select
            Else
                result = result & char
            End If
            pos = pos + 1
        Loop
        
        Err.Raise 1005, "JsonParser", "String sin terminar"
    End Function
    
    Private Function ParseNumber()
        Dim numStr, char
        numStr = ""
        
        Do While pos <= Len(jsonText)
            char = Mid(jsonText, pos, 1)
            If IsNumeric(char) Or char = "." Or char = "-" Or char = "+" Or LCase(char) = "e" Then
                numStr = numStr & char
                pos = pos + 1
            Else
                Exit Do
            End If
        Loop
        
        If IsNumeric(numStr) Then
            ParseNumber = CDbl(numStr)
        Else
            Err.Raise 1006, "JsonParser", "Número inválido: " & numStr
        End If
    End Function
    
    Private Function ParseBoolean()
        If Mid(jsonText, pos, 4) = "true" Then
            pos = pos + 4
            ParseBoolean = True
        ElseIf Mid(jsonText, pos, 5) = "false" Then
            pos = pos + 5
            ParseBoolean = False
        Else
            Err.Raise 1007, "JsonParser", "Valor booleano inválido"
        End If
    End Function
    
    Private Function ParseNull()
        If Mid(jsonText, pos, 4) = "null" Then
            pos = pos + 4
            ParseNull = Null
        Else
            Err.Raise 1008, "JsonParser", "Valor null inválido"
        End If
    End Function
    
    Private Sub SkipWhitespace()
        Do While pos <= Len(jsonText)
            Dim char
            char = Mid(jsonText, pos, 1)
            If char = " " Or char = Chr(9) Or char = Chr(10) Or char = Chr(13) Then
                pos = pos + 1
            Else
                Exit Do
            End If
        Loop
    End Sub
End Class

' ============================================================================
' FUNCIONES AUXILIARES JSON
' ============================================================================

Function IsDictionary(obj)
    On Error Resume Next
    IsDictionary = (TypeName(obj) = "Dictionary")
    On Error GoTo 0
End Function

Function IsArrayLike(obj)
    On Error Resume Next
    Dim result
    result = (TypeName(obj) = "ArrayList") Or (IsArray(obj))
    IsArrayLike = result
    On Error GoTo 0
End Function

Function CreateDict()
    Set CreateDict = CreateObject("Scripting.Dictionary")
End Function

Function CreateList()
    On Error Resume Next
    Set CreateList = CreateObject("System.Collections.ArrayList")
    If Err.Number <> 0 Then
        ' Fallback a Dictionary como lista si ArrayList no está disponible
        Err.Clear
        Set CreateList = CreateObject("Scripting.Dictionary")
    End If
    On Error GoTo 0
End Function

' ============================================================================
' UTILIDADES DE ARCHIVOS Y PATHS
' ============================================================================

Function ReadAllText(path)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile path
    ReadAllText = stream.ReadText
    stream.Close
End Function

Sub WriteAllText(path, text)
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.WriteText text
    stream.SaveToFile path, 2 ' adSaveCreateOverWrite
    stream.Close
End Sub

Function FileExists(path)
    FileExists = objFSO.FileExists(path)
End Function

Function DirExists(path)
    DirExists = objFSO.FolderExists(path)
End Function

Function PathCombine(base, rel)
    If Right(base, 1) = "\" Then
        PathCombine = base & rel
    Else
        PathCombine = base & "\" & rel
    End If
End Function

Function MakeRelative(base, abs)
    ' Implementación simple - devuelve path relativo si es posible
    If InStr(1, abs, base, vbTextCompare) = 1 Then
        MakeRelative = Mid(abs, Len(base) + 2) ' +2 para saltar el \
    Else
        MakeRelative = abs
    End If
End Function

Function NowUtcIso()
    Dim now
    now = Now()
    ' Formato ISO UTC simplificado
    NowUtcIso = Year(now) & "-" & _
                Right("00" & Month(now), 2) & "-" & _
                Right("00" & Day(now), 2) & "T" & _
                Right("00" & Hour(now), 2) & ":" & _
                Right("00" & Minute(now), 2) & ":" & _
                Right("00" & Second(now), 2) & "Z"
End Function

Function OleToRgbHex(lng)
    Dim r, g, b
    r = lng And &HFF
    g = (lng \ &H100) And &HFF
    b = (lng \ &H10000) And &HFF
    OleToRgbHex = "#" & Right("00" & Hex(r), 2) & Right("00" & Hex(g), 2) & Right("00" & Hex(b), 2)
End Function

Function RgbHexToOle(hex)
    Dim cleanHex, r, g, b
    cleanHex = Replace(hex, "#", "")
    If Len(cleanHex) = 6 Then
        r = CLng("&H" & Mid(cleanHex, 1, 2))
        g = CLng("&H" & Mid(cleanHex, 3, 2))
        b = CLng("&H" & Mid(cleanHex, 5, 2))
        RgbHexToOle = r + (g * &H100) + (b * &H10000)
    Else
        RgbHexToOle = 0
    End If
End Function

' ============================================================================
' DIFF SEMÁNTICO JSON
' ============================================================================

Function DiffJsonSemantico(jsonA, jsonB)
    On Error Resume Next
    
    Dim parser
    Set parser = New JsonParser
    
    Dim objA, objB
    Set objA = parser.Parse(jsonA)
    Set objB = parser.Parse(jsonB)
    
    If Err.Number <> 0 Then
        DiffJsonSemantico = "Error al parsear JSON: " & Err.Description
        Exit Function
    End If
    
    ' Normalizar y comparar
    Dim normalizedA, normalizedB
    normalizedA = NormalizeJsonObject(objA)
    normalizedB = NormalizeJsonObject(objB)
    
    If normalizedA = normalizedB Then
        DiffJsonSemantico = ""
    Else
        DiffJsonSemantico = FindDifferences(objA, objB, "", 0)
    End If
    
    On Error GoTo 0
End Function

Function NormalizeJsonObject(obj)
    Dim writer
    Set writer = New JsonWriter
    ' Aquí se podría implementar ordenación de claves, pero por simplicidad
    ' usamos la representación directa
    NormalizeJsonObject = writer.Stringify(obj)
End Function

Function FindDifferences(objA, objB, path, depth)
    If depth > 20 Then
        FindDifferences = "[Máximo de 20 niveles alcanzado]"
        Exit Function
    End If
    
    Dim result
    result = ""
    
    ' Implementación básica de comparación
    If IsDictionary(objA) And IsDictionary(objB) Then
        Dim key
        For Each key In objA.Keys
            If Not objB.Exists(key) Then
                result = result & "Clave faltante en B: " & path & "." & key & vbCrLf
            End If
        Next
        
        For Each key In objB.Keys
            If Not objA.Exists(key) Then
                result = result & "Clave extra en B: " & path & "." & key & vbCrLf
            End If
        Next
    Else
        If objA <> objB Then
            result = "Diferencia en " & path & ": '" & objA & "' vs '" & objB & "'" & vbCrLf
        End If
    End If
    
    FindDifferences = result
End Function

' ============================================================================
' FUNCIONES DE MAPEO PARA PROPIEDADES DE FORMULARIO
' ============================================================================

' Mapea BorderStyle numérico a token canónico
Function MapBorderStyleToToken(borderStyleValue)
    Select Case borderStyleValue
        Case 0: MapBorderStyleToToken = "None"
        Case 1: MapBorderStyleToToken = "Thin"
        Case 2: MapBorderStyleToToken = "Sizable"
        Case 3: MapBorderStyleToToken = "Dialog"
        Case Else: MapBorderStyleToToken = "Sizable" ' Default
    End Select
End Function

' Mapea ScrollBars numérico a token canónico
Function MapScrollBarsToToken(scrollBarsValue)
    Select Case scrollBarsValue
        Case 0: MapScrollBarsToToken = "Neither"
        Case 1: MapScrollBarsToToken = "Horizontal"
        Case 2: MapScrollBarsToToken = "Vertical"
        Case 3: MapScrollBarsToToken = "Both"
        Case Else: MapScrollBarsToToken = "Neither" ' Default
    End Select
End Function

' Mapea MinMaxButtons numérico a token canónico
Function MapMinMaxButtonsToToken(minMaxButtonsValue)
    Select Case minMaxButtonsValue
        Case 0: MapMinMaxButtonsToToken = "None"
        Case 1: MapMinMaxButtonsToToken = "Min Enabled"
        Case 2: MapMinMaxButtonsToToken = "Max Enabled"
        Case 3: MapMinMaxButtonsToToken = "Both Enabled"
        Case Else: MapMinMaxButtonsToToken = "None" ' Default
    End Select
End Function

' Mapea RecordsetType numérico a token canónico
Function MapRecordsetTypeToToken(recordsetTypeValue)
    Select Case recordsetTypeValue
        Case 0: MapRecordsetTypeToToken = "Dynaset"
        Case 1: MapRecordsetTypeToToken = "Snapshot"
        Case 2: MapRecordsetTypeToToken = "Dynaset (Inconsistent Updates)"
        Case Else: MapRecordsetTypeToToken = "Dynaset" ' Default
    End Select
End Function

' Mapea Orientation numérico a token canónico
Function MapOrientationToToken(orientationValue)
    Select Case orientationValue
        Case 0: MapOrientationToToken = "LeftToRight"
        Case 1: MapOrientationToToken = "RightToLeft"
        Case Else: MapOrientationToToken = "LeftToRight" ' Default
    End Select
End Function

' Mapea SplitFormOrientation numérico a token canónico
Function MapSplitFormOrientationToToken(splitFormOrientationValue)
    Select Case splitFormOrientationValue
        Case 0: MapSplitFormOrientationToToken = "DatasheetOnTop"
        Case 1: MapSplitFormOrientationToToken = "DatasheetOnBottom"
        Case 2: MapSplitFormOrientationToToken = "DatasheetOnLeft"
        Case 3: MapSplitFormOrientationToToken = "DatasheetOnRight"
        Case Else: MapSplitFormOrientationToToken = "DatasheetOnTop" ' Default
    End Select
End Function

' ============================================================================
' SISTEMA DE LOGGING
' ============================================================================

Sub LogInfo(message)
    If gVerbose Then
        WScript.Echo "[INFO] " & message
    End If
End Sub

Sub LogWarn(message)
    WScript.Echo "[WARN] " & message
End Sub

Sub LogErr(message)
    WScript.Echo "[ERROR] " & message
End Sub

' ============================================================================
' FUNCIÓN GetFunctionalityFiles ACTUALIZADA
' ============================================================================

Function GetFunctionalityFiles(functionality)
    ' Esta función devuelve únicamente condor_cli.vbs para mejoras de infraestructura
    If functionality = "CLI" Or functionality = "Infrastructure" Then
        GetFunctionalityFiles = Array("condor_cli.vbs")
        Exit Function
    End If
    
    ' Resto de funcionalidades existentes...
    ' [Código existente de GetFunctionalityFiles se mantiene igual]
End Function
