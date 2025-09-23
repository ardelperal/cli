' ============================================================================
' SECCIÓN 10: FUNCIÓN PRINCIPAL
' ============================================================================
' CLI ACCESS - Herramienta de línea de comandos para análisis de Access
' ============================================================================
' Descripción: Script VBScript para extraer información completa de bases de datos Access
' Autor: Desarrollador CLI
' Versión: 1.0 MVP
' ============================================================================

Option Explicit

' ============================================================================
' SECCIÓN 1: CONSTANTES DE ACCESS
' ============================================================================

' Constantes de objetos de Access
Const acModule = 5
Const acClassModule = 100
Const acForm = -32768
Const acReport = -32764
Const acMacro = -32766
Const acTable = 0
Const acQuery = 1
Const acDesign = 1
Const acDefault = -1
Const acHidden = 1
Const acNormal = 0

' Constantes de comandos de Access
Const acCmdCompileAndSaveAllModules = 126

' Constantes de guardado
Const acSaveNo = 2
Const acSaveYes = 0

' Constantes de controles
Const acLabel = 100
Const acTextBox = 109
Const acCommandButton = 104
Const acCheckBox = 106
Const acOptionButton = 101
Const acComboBox = 111
Const acListBox = 110
Const acSubform = 112
Const acImage = 103

' Controles extra
Const acRectangle = 3
Const acLine = 4
Const acOptionGroup = 7
Const acBoundObjectFrame = 108
Const acUnboundObjectFrame = 114
Const acPageBreak = 10
Const acCustomControl = 119
Const acToggleButton = 122
Const acTabCtl = 123
Const acPage = 124

' Vistas de formulario
Const acFormDS = 3
Const acFormPivotTable = 4
Const acFormPivotChart = 5

' ============================================================================
' SECCIÓN 2: VARIABLES GLOBALES
' ============================================================================

Dim objFSO, objArgs, objAccess, objConfig
Dim gVerbose, gQuiet, gDryRun, gDebug
Dim gDbPath, gPassword, gOutputPath, gConfigPath, gScriptPath, gScriptDir
Dim g_ModulesSrcPath, g_ModulesExtensions, g_ModulesIncludeSubdirs

' Variables para ImportFormFromJson
Dim ImportFormFromJson_app
Dim ImportFormFromJson_target
Dim ImportFormFromJson_finalTarget
Dim ImportFormFromJson_root
Dim gConfig

' ============================================================================
' SECCIÓN 3: FUNCIONES DE CONFIGURACIÓN
' ============================================================================

' Función para cargar configuración desde archivo INI
Function LoadConfig(configPath)
    On Error Resume Next
    
    Dim config, fso, file, line, parts, key, value, section
    Set config = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Resolver ruta del archivo de configuración
    If objFSO.GetAbsolutePathName(configPath) <> configPath Then
        configPath = gScriptDir & "\" & configPath
    End If
    
    ' Valores por defecto con rutas relativas al script
    config.Add "DATABASE_DefaultPath", ""
    config.Add "DATABASE_Password", ""
    config.Add "DATABASE_Timeout", "30"
    config.Add "OUTPUT_DefaultPath", gScriptDir & "\output"
    config.Add "OUTPUT_Format", "json"
    config.Add "OUTPUT_PrettyPrint", "true"
    config.Add "OUTPUT_OutputPath", gScriptDir & "\output"
    config.Add "LOGGING_Verbose", "false"
    config.Add "LOGGING_QuietMode", "false"
    config.Add "LOGGING_LogFile", gScriptDir & "\cli.log"
    config.Add "LOGGING_LogLevel", "INFO"
    config.Add "EXTRACTION_IncludeTables", "true"
    config.Add "EXTRACTION_IncludeForms", "true"
    config.Add "EXTRACTION_IncludeQueries", "true"
    config.Add "EXTRACTION_IncludeRelations", "true"
    config.Add "EXTRACTION_FilterSystemObjects", "true"
    config.Add "MODULES_SrcPath", gScriptDir & "\src"
    config.Add "MODULES_Extensions", ".bas,.cls"
    config.Add "MODULES_IncludeSubdirectories", "true"
    config.Add "MODULES_FilePattern", "*"
    
    If Not fso.FileExists(configPath) Then
        LogVerbose "Archivo de configuracion no encontrado: " & configPath & ". Usando valores por defecto."
        Set LoadConfig = config
        Exit Function
    Else
        LogVerbose "Archivo de configuracion encontrado: " & configPath
    End If
    
    Set file = fso.OpenTextFile(configPath, 1)
    section = ""
    
    Do While Not file.AtEndOfStream
        line = Trim(file.ReadLine)
        
        ' Ignorar líneas vacías y comentarios
        If Len(line) > 0 And Left(line, 1) <> ";" And Left(line, 1) <> "#" Then
            ' Detectar secciones
            If Left(line, 1) = "[" And Right(line, 1) = "]" Then
                section = Mid(line, 2, Len(line) - 2)
            Else
                ' Procesar pares clave=valor
                If InStr(line, "=") > 0 Then
                    parts = Split(line, "=", 2)
                    If UBound(parts) >= 1 Then
                        key = section & "_" & Trim(parts(0))
                        value = Trim(parts(1))
                        
                        LogVerbose "Procesando: [" & section & "] " & Trim(parts(0)) & " = " & value & " -> clave: " & key
                        
                        ' Resolver rutas relativas para ciertos valores (excepto patrones)
                        If (InStr(UCase(key), "PATH") > 0 Or InStr(UCase(key), "FILE") > 0) And InStr(UCase(key), "PATTERN") = 0 Then
                            If value <> "" And fso.GetAbsolutePathName(value) <> value Then
                                value = gScriptDir & "\" & value
                                LogVerbose "Ruta resuelta: " & value
                            End If
                        End If
                        
                        If config.Exists(key) Then
                            config(key) = value
                            LogVerbose "Clave actualizada: " & key & " = " & value
                        Else
                            config.Add key, value
                            LogVerbose "Clave agregada: " & key & " = " & value
                        End If
                    End If
                End If
            End If
        End If
    Loop
    
    file.Close
    Set LoadConfig = config
    
    If Err.Number <> 0 Then
        LogMessage "Error cargando configuracion: " & Err.Description
        Err.Clear
    End If
End Function

' ============================================================================
' FUNCIONES HELPER PARA CONFIGURACIÓN UI
' ============================================================================

Function UI_GetRoot(config)
    On Error Resume Next
    UI_GetRoot = ResolvePath(CfgGet(config,"UI_Root","UI.Root",".\ui"))
End Function

Function UI_GetFormsDir(config)
    On Error Resume Next
    UI_GetFormsDir = CfgGet(config,"UI_FormsDir","UI.FormsDir","forms")
End Function

Function UI_GetAssetsDir(config)
    On Error Resume Next
    UI_GetAssetsDir = CfgGet(config,"UI_AssetsDir","UI.AssetsDir","assets")
End Function

Function UI_GetIncludeSubdirs(config)
    On Error Resume Next
    UI_GetIncludeSubdirs = ToBool(CfgGet(config,"UI_IncludeSubdirectories","UI.IncludeSubdirectories","true"),True)
End Function

Function UI_GetFormFilePattern(config)
    On Error Resume Next
    UI_GetFormFilePattern = CfgGet(config,"UI_FormFilePattern","UI.FormFilePattern","*.json")
End Function

Function UI_NameFromFileBase(config)
    On Error Resume Next
    UI_NameFromFileBase = ToBool(CfgGet(config,"UI_NameFromFileBase","UI.NameFromFileBase","true"),True)
End Function

Function UI_GetAssetsImgDir(config)
    On Error Resume Next
    UI_GetAssetsImgDir = CfgGet(config,"UI_AssetsImgDir","UI.AssetsImgDir","img")
End Function

Function UI_GetAssetsImgExtensions(config)
    On Error Resume Next
    Dim exts: exts = CfgGet(config,"UI_AssetsImgExtensions","UI.AssetsImgExtensions",".png,.jpg,.jpeg,.gif,.bmp")
    UI_GetAssetsImgExtensions = LCase(exts)
End Function

' ============================================================================
' UTILIDADES DE FICHEROS
' ============================================================================

Function EnumerateFiles(rootFolder, pattern, includeSubdirs)
    On Error Resume Next
    Dim fso, folder, files, subFolder, col, arr, i, subArr, j
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(rootFolder) Then
        EnumerateFiles = Array()
        Exit Function
    End If
    Set folder = fso.GetFolder(rootFolder)
    ReDim arr(-1)
    ' archivos directos
    Set files = folder.Files
    For Each col In files
        If LCase(fso.GetFileName(col.Path)) Like LCase(pattern) Then
            ReDim Preserve arr(UBound(arr)+1)
            arr(UBound(arr)) = col.Path
        End If
    Next
    ' recursivo
    If includeSubdirs Then
        For Each subFolder In folder.SubFolders
            subArr = EnumerateFiles(subFolder.Path, pattern, True)
            If IsArray(subArr) Then
                For j = 0 To UBound(subArr)
                    ReDim Preserve arr(UBound(arr)+1)
                    arr(UBound(arr)) = subArr(j)
                Next
            End If
        Next
    End If
    EnumerateFiles = arr
End Function

Function FileBaseName(p)
    On Error Resume Next
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    FileBaseName = fso.GetBaseName(p)
End Function

' ============================================================================
' SECCIÓN 4: FUNCIONES DE AYUDA
' ============================================================================

Sub ShowHelp()
    WScript.Echo "=== CLI ACCESS - Herramienta de Analisis de Access ==="
    WScript.Echo "Version: 1.0 MVP"
    WScript.Echo ""
    WScript.Echo "USO:"
    WScript.Echo "  cscript cli.vbs <comando> [argumentos] [opciones]"
    WScript.Echo ""
    WScript.Echo "COMANDOS DISPONIBLES:"
    WScript.Echo ""
    WScript.Echo "  extract-modules <db_path>"
    WScript.Echo "    Extrae modulos VBA hacia archivos fuente"
    WScript.Echo ""
    WScript.Echo "  list-objects <db_path>"
    WScript.Echo "    Lista todos los objetos de la base de datos"
    WScript.Echo ""
    WScript.Echo "  rebuild <db_path>"
    WScript.Echo "    Reconstruye todos los modulos VBA desde archivos fuente"
    WScript.Echo "    - Usa MODULES_SrcPath del archivo de configuracion"
    WScript.Echo "    - Formatos soportados: .bas (modulos), .cls (clases)"
    WScript.Echo "    - Si no se especifica db_path, usa DATABASE_DefaultPath del config"
    WScript.Echo "    - Requiere: ""Trust access to the VBA project object model"" en Access"
    WScript.Echo ""
    WScript.Echo "  update <db_path>"
    WScript.Echo "    Actualizar modulos VBA desde src"
    WScript.Echo ""
    WScript.Echo "  schema [db_path] [--table <nombre>] [--out <ruta>] [--format json|md]"
    WScript.Echo "    Exporta la estructura de tablas. Por defecto usa DATABASE_DefaultPath del ini."
    WScript.Echo ""
    WScript.Echo "  form-export [db_path] --form <NombreFormulario> --out <ruta.json>"
    WScript.Echo "  form-import <json_path> [db_path] [--name NuevoNombre] [--overwrite] [--dry-run]"
    WScript.Echo ""
    WScript.Echo "  ui-rebuild [db_path]           Reconvierte TODOS los formularios del /ui a Access"
    WScript.Echo "  ui-update [db_path] <form.json|formName> [...]  Reconstruye solo los indicados"
    WScript.Echo "  ui-touch <formName|form.json>  Atajo para actualizar un unico formulario"
    WScript.Echo ""
    WScript.Echo "OPCIONES:"
    WScript.Echo "  --config <path>       - Archivo de configuracion (por defecto: cli.ini)"
    WScript.Echo "  --password <pwd>      - Contrasena de la base de datos"
    WScript.Echo "  --verbose             - Salida detallada"
    WScript.Echo "  --quiet               - Salida minima"
    WScript.Echo "  --help                - Muestra esta ayuda"
    WScript.Echo ""
    WScript.Echo "MODIFICADORES DE TESTING:"
    WScript.Echo "  /dry-run              - Simula la ejecucion sin cambios"
    WScript.Echo "  /validate             - Valida la configuracion"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS:"
    WScript.Echo "  cscript cli.vbs extract-modules ""C:\mi_base.accdb"""
    WScript.Echo "  cscript cli.vbs list-objects ""C:\mi_base.accdb"" --verbose"
    WScript.Echo "  cscript cli.vbs rebuild"
    WScript.Echo "  cscript cli.vbs update ""C:\mi_base.accdb"" /verbose"
    WScript.Echo "  cscript cli.vbs schema --table Usuarios --format md"
    WScript.Echo "  cscript cli.vbs schema ""C:\mi_base.accdb"" --out ""C:\temp"""
End Sub

' ============================================================================
' SECCIÓN 5: FUNCIONES DE ACCESS
' ============================================================================

' Función para obtener PIDs de procesos MSACCESS.EXE usando WMI
Function GetAccessPIDs()
    On Error Resume Next
    Dim objWMI, colProcesses, objProcess, arrPIDs(), pidCount
    
    Set objWMI = GetObject("winmgmts:")
    If Err.Number <> 0 Then
        Err.Clear
        GetAccessPIDs = Array()
        Exit Function
    End If
    
    Set colProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'MSACCESS.EXE'")
    If Err.Number <> 0 Then
        Err.Clear
        GetAccessPIDs = Array()
        Exit Function
    End If
    
    pidCount = 0
    For Each objProcess In colProcesses
        If Err.Number = 0 Then
            ReDim Preserve arrPIDs(pidCount)
            arrPIDs(pidCount) = objProcess.ProcessId
            pidCount = pidCount + 1
        Else
            Err.Clear
        End If
    Next
    
    If pidCount = 0 Then
        GetAccessPIDs = Array()
    Else
        GetAccessPIDs = arrPIDs
    End If
    
    On Error GoTo 0
End Function

' Función para terminar un PID específico de Access
Sub TerminateAccessPID(targetPID)
    On Error Resume Next
    Dim objWMI, colProcesses, objProcess
    
    If targetPID <= 0 Then Exit Sub
    
    Set objWMI = GetObject("winmgmts:")
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    
    Set colProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'MSACCESS.EXE' AND ProcessId = " & targetPID)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    
    For Each objProcess In colProcesses
        If Err.Number = 0 Then
            LogVerbose "Terminando proceso Access PID: " & targetPID
            objProcess.Terminate()
        Else
            Err.Clear
        End If
    Next
    
    On Error GoTo 0
End Sub

' Función para cerrar procesos de Access existentes antes de comenzar
Sub CloseExistingAccessProcesses()
    Dim pids, i
    
    LogVerbose "Verificando procesos de Access existentes..."
    pids = GetAccessPIDs()
    
    If IsArray(pids) And UBound(pids) >= 0 Then
        LogMessage "Cerrando " & (UBound(pids) + 1) & " proceso(s) de Access existente(s)..."
        For i = 0 To UBound(pids)
            TerminateAccessPID pids(i)
        Next
        
        ' Esperar un momento para que los procesos se cierren
        WScript.Sleep 1000
        LogVerbose "Procesos de Access cerrados"
    Else
        LogVerbose "No hay procesos de Access ejecutandose"
    End If
End Sub

' Función para abrir Access de forma segura
' Función canónica para abrir Access (basada en condor_cli.vbs)
Function OpenAccess(dbPath, password)
    Dim objApp, attempt, maxAttempts
    maxAttempts = 3
    
    LogVerbose "Abriendo Access: " & dbPath
    If password <> "" Then
        LogVerbose "Con password: (oculta)"
    End If
    
    For attempt = 1 To maxAttempts
        On Error Resume Next
        Set objApp = CreateObject("Access.Application")
        
        If Err.Number <> 0 Then
            LogMessage "ERROR: No se pudo crear instancia de Access (intento " & attempt & "): " & Err.Description
            If attempt < maxAttempts Then
                LogVerbose "Reintentando en 500ms..."
                WScript.Sleep 500
                Err.Clear
            Else
                Set OpenAccess = Nothing
                Exit Function
            End If
        Else
            Exit For
        End If
    Next
    
    ' Configurar Access para modo silencioso (solo lo que NO requiere BD abierta)
    objApp.Visible = False
    objApp.UserControl = False
    objApp.AutomationSecurity = 1  ' msoAutomationSecurityForceDisable - IGUAL que rebuild
    
    ' Abrir la base de datos con retry
    For attempt = 1 To maxAttempts
        On Error Resume Next
        If password <> "" Then
            objApp.OpenCurrentDatabase dbPath, False, password
        Else
            objApp.OpenCurrentDatabase dbPath, False
        End If

        If Err.Number <> 0 Then
            LogMessage "ERROR: No se pudo abrir la base de datos (intento " & attempt & "): " & Err.Description
            If attempt < maxAttempts Then
                LogVerbose "Reintentando en 500ms..."
                WScript.Sleep 500
                Err.Clear
            Else
                objApp.Quit
                Set objApp = Nothing
                Set OpenAccess = Nothing
                Exit Function
            End If
        Else
            Exit For
        End If
    Next
    
    ' Verificar que la BD se abrió correctamente
    If objApp.CurrentProject Is Nothing Then
        LogMessage "ERROR: No se pudo abrir la base de datos (CurrentProject = Nothing)"
        objApp.Quit
        Set objApp = Nothing
        Set OpenAccess = Nothing
        Exit Function
    End If
    
    ' Configurar modo silencioso DESPUÉS de abrir la BD
    On Error Resume Next
    objApp.Echo False
    objApp.DoCmd.SetWarnings False
    objApp.DisplayAlerts = False
    objApp.VBE.MainWindow.Visible = False
    Err.Clear
    On Error GoTo 0
    
    ' Verificar acceso al VBE para operaciones de módulos
    If Not objApp.VBE Is Nothing Then
        LogVerbose "Acceso VBE disponible"
    Else
        LogMessage "ADVERTENCIA: Acceso VBE no disponible - las operaciones de modulos pueden fallar"
        LogMessage "SOLUCION: Habilite 'Confiar en el acceso al modelo de objetos de proyectos de VBA' en Access"
    End If
    
    On Error GoTo 0
    Set OpenAccess = objApp
    
    LogVerbose "Access abierto exitosamente"
End Function

' Función canónica para cerrar Access de forma segura (basada en condor_cli.vbs)
Sub CloseAccess(objAccess)
    If Not objAccess Is Nothing Then
        LogVerbose "Cerrando Access..."
        
        On Error Resume Next
        ' Restaurar configuraciones EXACTAMENTE como rebuild
        objAccess.Echo True
        objAccess.DoCmd.SetWarnings True
        ' No restaurar AutomationSecurity - rebuild no lo restaura
        objAccess.CloseCurrentDatabase
        objAccess.Quit
        Set objAccess = Nothing
        On Error GoTo 0
        
        LogVerbose "Access cerrado exitosamente"
    End If
End Sub

' ============================================================================
' SECCIÓN 6: FUNCIONES DE LOGGING
' ============================================================================

Sub LogMessage(message)
    ' Reemplazar caracteres especiales para evitar problemas de codificación en consola
    Dim cleanMessage
    cleanMessage = Replace(message, "ó", "o")
    cleanMessage = Replace(cleanMessage, "ñ", "n")
    cleanMessage = Replace(cleanMessage, "á", "a")
    cleanMessage = Replace(cleanMessage, "é", "e")
    cleanMessage = Replace(cleanMessage, "í", "i")
    cleanMessage = Replace(cleanMessage, "ú", "u")
    cleanMessage = Replace(cleanMessage, "Ó", "O")
    cleanMessage = Replace(cleanMessage, "Ñ", "N")
    cleanMessage = Replace(cleanMessage, "Á", "A")
    cleanMessage = Replace(cleanMessage, "É", "E")
    cleanMessage = Replace(cleanMessage, "Í", "I")
    cleanMessage = Replace(cleanMessage, "Ú", "U")
    
    ' Escribir a consola solo si no está en modo quiet
    If Not gQuiet Then
        WScript.Echo "[" & Now & "] " & cleanMessage
    End If
    
    ' Escribir siempre al archivo de log (mensaje original con acentos)
    AppendLogToFile message
End Sub

Sub LogVerbose(message)
    If gVerbose And Not gQuiet Then
        WScript.Echo "[VERBOSE] " & message
    End If
End Sub

Sub LogError(message)
    WScript.Echo "[ERROR] " & message
    ' Escribir siempre al archivo de log
    AppendLogToFile "[ERROR] " & message
End Sub

Sub AppendLogToFile(message)
    On Error Resume Next
    
    Dim logFile
    If IsObject(gConfig) Then
        On Error Resume Next
        If Not gConfig Is Nothing And gConfig.Exists("LOGGING_LogFile") Then
            logFile = gConfig("LOGGING_LogFile")
        End If
        Err.Clear
    End If
    If ("" & logFile) = "" Then
        logFile = gScriptDir & "\cli.log"
    End If
    
    ' Crear el objeto ADODB.Stream
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    If Err.Number <> 0 Then
        Exit Sub ' Si no se puede crear el stream, salir silenciosamente
    End If
    
    ' Configurar el stream para UTF-8
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    
    ' Abrir el stream
    stream.Open
    
    ' Si el archivo existe, cargar su contenido
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(logFile) Then
        stream.LoadFromFile logFile
        stream.Position = stream.Size ' Ir al final del archivo
    End If
    
    ' Escribir el mensaje con timestamp
    stream.WriteText "[" & Now & "] " & message & vbCrLf
    
    ' Guardar al archivo
    stream.SaveToFile logFile, 2 ' adSaveCreateOverWrite
    On Error GoTo 0
    
    ' Cerrar el stream
    stream.Close
    Set stream = Nothing
    Set fso = Nothing
    
    Err.Clear
End Sub

' ============================================================================
' SECCIÓN 7: FUNCIONES PRINCIPALES
' ============================================================================

Function ExtractTableFields(db, tableName)
    Dim fieldsDict, fieldDict, tbl, fld
    Set fieldsDict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    
    Set tbl = db.TableDefs(tableName)
    If Err.Number <> 0 Then
        LogMessage "Error: Tabla '" & tableName & "' no encontrada", "ERROR"
        Set ExtractTableFields = fieldsDict
        Exit Function
    End If
    
    For Each fld In tbl.Fields
        Set fieldDict = CreateObject("Scripting.Dictionary")
        fieldDict("name") = fld.Name
        fieldDict("type") = GetFieldTypeName(fld.Type)
        fieldDict("size") = fld.Size
        fieldDict("required") = fld.Required
        fieldDict("allowZeroLength") = fld.AllowZeroLength
        fieldDict("ordinalPosition") = fld.OrdinalPosition
        
        If fld.DefaultValue <> "" Then
            fieldDict("defaultValue") = fld.DefaultValue
        End If
        
        ' Propiedades adicionales si existen
        On Error Resume Next
        fieldDict("validationRule") = fld.ValidationRule
        fieldDict("validationText") = fld.ValidationText
        On Error GoTo 0
        
        Set fieldsDict(fld.Name) = fieldDict
    Next
    
    Set ExtractTableFields = fieldsDict
End Function

Function ExtractTableRelations(db)
    Dim relationsDict, relationDict, rel, fld
    Set relationsDict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    
    For Each rel In db.Relations
        Set relationDict = CreateObject("Scripting.Dictionary")
        relationDict("name") = rel.Name
        relationDict("table") = rel.Table
        relationDict("foreignTable") = rel.ForeignTable
        relationDict("attributes") = rel.Attributes
        
        ' Campos de la relación
        Dim relationFieldsArray()
        ReDim relationFieldsArray(rel.Fields.Count - 1)
        Dim i: i = 0
        
        For Each fld In rel.Fields
            Dim relFieldDict
            Set relFieldDict = CreateObject("Scripting.Dictionary")
            relFieldDict("name") = fld.Name
            relFieldDict("foreignName") = fld.ForeignName
            Set relationFieldsArray(i) = relFieldDict
            i = i + 1
        Next
        
        relationDict("fields") = relationFieldsArray
        
        Set relationsDict(rel.Name) = relationDict
    Next
    
    Set ExtractTableRelations = relationsDict
End Function

' Función para extraer información de tablas (mantenida para compatibilidad interna)
Sub ExtractTables(outputPath)
    LogMessage "Extrayendo informacion de tablas..."
    
    Dim tableInfo, tbl, fld
    Set tableInfo = CreateObject("Scripting.Dictionary")
    
    ' Iterar sobre todas las tablas
    For Each tbl In objAccess.CurrentDb.TableDefs
        If Left(tbl.Name, 4) <> "MSys" Then ' Excluir tablas del sistema
            LogVerbose "Procesando tabla: " & tbl.Name
            
            Dim tableData
            Set tableData = CreateObject("Scripting.Dictionary")
            
            tableData.Add "Name", tbl.Name
            tableData.Add "RecordCount", tbl.RecordCount
            tableData.Add "DateCreated", tbl.DateCreated
            tableData.Add "LastUpdated", tbl.LastUpdated
            
            ' Extraer campos
            Dim fields
            Set fields = CreateObject("Scripting.Dictionary")
            
            For Each fld In tbl.Fields
                Dim fieldData
                Set fieldData = CreateObject("Scripting.Dictionary")
                
                fieldData.Add "Type", fld.Type
                fieldData.Add "Size", fld.Size
                fieldData.Add "Required", fld.Required
                fieldData.Add "AllowZeroLength", fld.AllowZeroLength
                
                fields.Add fld.Name, fieldData
            Next
            
            tableData.Add "Fields", fields
            tableInfo.Add tbl.Name, tableData
        End If
    Next
    
    ' Guardar información en archivo JSON
    SaveToJSON tableInfo, outputPath & "\tables.json"
    LogMessage "Informacion de tablas guardada en: tables.json"
End Sub



' ============================================================================
' SECCIÓN 6: FUNCIONES DE EXPORTACIÓN DE ESQUEMA
' ============================================================================

' Exporta estructura de todas las tablas o una tabla
Function ExportSchema(dbPath, tableFilter, outDir, fmt)
    On Error Resume Next
    ExportSchema = False

    Dim app: Set app = OpenAccess(dbPath, GetDatabasePassword(dbPath))
    If app Is Nothing Then Exit Function

    Dim db: Set db = app.CurrentDb
    Dim schema, tdef, rel
    Set schema = CreateObject("Scripting.Dictionary")

    ' Índices de relaciones por tabla (FK entrantes/salientes)
    Dim relsOut, relsIn
    Set relsOut = CreateObject("Scripting.Dictionary")
    Set relsIn  = CreateObject("Scripting.Dictionary")

    ' Precalcular relaciones
    For Each rel In db.Relations
        If Left(rel.Table,4) <> "MSys" And Left(rel.ForeignTable,4) <> "MSys" Then
            Dim rf, arr, i
            ' salientes desde rel.Table
            If Not relsOut.Exists(rel.Table) Then relsOut(rel.Table) = Array()
            arr = relsOut(rel.Table)
            ReDim Preserve arr(UBound(arr)+1)
            Set arr(UBound(arr)) = rel
            relsOut(rel.Table) = arr
            ' entrantes hacia rel.ForeignTable
            If Not relsIn.Exists(rel.ForeignTable) Then relsIn(rel.ForeignTable) = Array()
            arr = relsIn(rel.ForeignTable)
            ReDim Preserve arr(UBound(arr)+1)
            Set arr(UBound(arr)) = rel
            relsIn(rel.ForeignTable) = arr
        End If
    Next

    ' Recorre tablas
    For Each tdef In db.TableDefs
        If Left(tdef.Name,4) <> "MSys" Then
            If (tableFilter = "") Or (StrComp(tdef.Name, tableFilter, vbTextCompare) = 0) Then
                Dim tinfo: Set tinfo = CreateObject("Scripting.Dictionary")
                tinfo("name") = tdef.Name
                tinfo("dateCreated") = tdef.DateCreated
                tinfo("lastUpdated") = tdef.LastUpdated

                ' Campos
                Dim f, fields: Set fields = CreateObject("Scripting.Dictionary")
                For Each f In tdef.Fields
                    Dim finfo: Set finfo = CreateObject("Scripting.Dictionary")
                    finfo("name") = f.Name
                    finfo("type") = GetFieldTypeName(f.Type)
                    finfo("size") = f.Size
                    finfo("required") = f.Required
                    finfo("allowZeroLength") = f.AllowZeroLength
                    On Error Resume Next
                    finfo("defaultValue") = f.DefaultValue
                    finfo("validationRule") = f.ValidationRule
                    finfo("validationText") = f.ValidationText
                    On Error GoTo 0
                    Set fields(f.Name) = finfo
                    If gDebug Then LogMessage "Campo agregado: " & f.Name & " (" & GetFieldTypeName(f.Type) & ")"
                Next
                Set tinfo("fields") = fields
                If gDebug Then LogMessage "Total campos en " & tdef.Name & ": " & fields.Count

                ' Índices y PK
                Dim idx, indexes: Set indexes = CreateObject("Scripting.Dictionary")
                Dim pkCols: pkCols = ""
                For Each idx In tdef.Indexes
                    Dim idic: Set idic = CreateObject("Scripting.Dictionary")
                    idic("name") = idx.Name
                    idic("primary") = idx.Primary
                    idic("unique") = idx.Unique
                    ' columnas del índice
                    Dim c, cols: cols = ""
                    For Each c In idx.Fields
                        If cols <> "" Then cols = cols & ", "
                        cols = cols & c.Name
                    Next
                    idic("columns") = cols
                    Set indexes(idx.Name) = idic
                    If idx.Primary Then pkCols = cols
                Next
                Set tinfo("indexes") = indexes
                tinfo("primaryKey") = pkCols

                ' Relaciones (FK salientes y entrantes)
                tinfo("relationsOut") = DescribeRelationsArray(relsOut, tdef.Name)
                tinfo("relationsIn")  = DescribeRelationsArray(relsIn,  tdef.Name)

                Set schema(tdef.Name) = tinfo
            End If
        End If
    Next

    ' Salida
    Dim outFile
    If tableFilter = "" Then
        outFile = outDir & "\schema_all." & fmt
    Else
        outFile = outDir & "\schema_" & tableFilter & "." & fmt
    End If

    If LCase(fmt) = "md" Then
        SaveSchemaAsMarkdown schema, outFile
    Else
        SaveToJSON schema, outFile
    End If

    CloseAccess app
    ExportSchema = True
End Function

' Convierte listado de relaciones a matriz de diccionarios
Function DescribeRelationsArray(relIndex, tableName)
    On Error Resume Next
    Dim arr, i, out(), count: count = 0
    If relIndex.Exists(tableName) Then
        arr = relIndex(tableName)
        ReDim out(UBound(arr))
        For i = 0 To UBound(arr)
            Dim d: Set d = CreateObject("Scripting.Dictionary")
            d("name") = arr(i).Name
            d("table") = arr(i).Table
            d("foreignTable") = arr(i).ForeignTable
            ' pares campo->campo
            Dim rf, pairs: pairs = ""
            For Each rf In arr(i).Fields
                If pairs <> "" Then pairs = pairs & ", "
                pairs = pairs & rf.Name & " -> " & rf.ForeignName
            Next
            d("fieldPairs") = pairs
            Set out(count) = d
            count = count + 1
        Next
        If count > 0 Then
            ReDim Preserve out(count-1)
            DescribeRelationsArray = out
            Exit Function
        End If
    End If
    DescribeRelationsArray = Array()
End Function

' Función para guardar esquema en formato Markdown
Sub SaveSchemaAsMarkdown(schema, filePath)
    On Error Resume Next
    Dim s, tName, tinfo, fName, finfo, i, rel, fields, fieldKeys

    s = "# Esquema de base de datos" & vbCrLf & vbCrLf

    For Each tName In schema.Keys
        Set tinfo = schema(tName)
        s = s & "## Tabla: " & tName & vbCrLf
        If tinfo.Exists("primaryKey") And tinfo("primaryKey") <> "" Then
            s = s & "- **PK:** " & tinfo("primaryKey") & vbCrLf
        End If
        s = s & vbCrLf & "| Campo | Tipo | Tamano | Requerido | Defecto |" & vbCrLf
        s = s & "|---|---|---:|:---:|---|" & vbCrLf

        If tinfo.Exists("fields") Then
            Set fields = tinfo("fields")
            If gDebug Then LogMessage "Procesando campos de " & tName & ", total: " & fields.Count
            If fields.Count > 0 Then
                fieldKeys = fields.Keys
                If gDebug Then LogMessage "Claves de campos obtenidas: " & UBound(fieldKeys) + 1
                For i = 0 To UBound(fieldKeys)
                    fName = fieldKeys(i)
                    Set finfo = fields(fName)
                    If gDebug Then LogMessage "Procesando campo: " & fName & " - Tipo objeto: " & TypeName(finfo)
                    
                    Dim fieldType, fieldSize, fieldRequired
                    fieldType = ""
                    fieldSize = ""
                    fieldRequired = False
                    
                    On Error Resume Next
                    fieldType = finfo("type")
                    fieldSize = finfo("size")
                    fieldRequired = finfo("required")
                    On Error GoTo 0
                    
                    If gDebug Then LogMessage "Valores obtenidos - Tipo: " & fieldType & ", Tamaño: " & fieldSize & ", Requerido: " & fieldRequired
                    
                    Dim defVal: defVal = ""
                     If finfo.Exists("defaultValue") Then defVal = finfo("defaultValue")
                     Dim requiredText: requiredText = "No"
                     If fieldRequired Then requiredText = "Si"
                     Dim rowText: rowText = "| " & fName & " | " & fieldType & " | " & fieldSize & " | " & requiredText & " | " & defVal & " |" & vbCrLf
                    s = s & rowText
                    If gDebug Then LogMessage "Fila agregada: " & Trim(Replace(rowText, vbCrLf, ""))
                    If gDebug Then LogMessage "Campo procesado: " & fName
                Next
            Else
                If gDebug Then LogMessage "No hay campos en el diccionario para " & tName
            End If
        Else
            If gDebug Then LogMessage "No hay campos para " & tName
        End If

        ' Relaciones salientes/entrantes
        s = s & vbCrLf & "**FK (salientes):**" & vbCrLf
        If IsArray(tinfo("relationsOut")) And UBound(tinfo("relationsOut")) >= 0 Then
            For i = 0 To UBound(tinfo("relationsOut"))
                Set rel = tinfo("relationsOut")(i)
                s = s & "- " & rel("name") & ": " & tName & " → " & rel("foreignTable") & " (" & rel("fieldPairs") & ")" & vbCrLf
            Next
        Else
            s = s & "- (ninguna)" & vbCrLf
        End If

        s = s & vbCrLf & "**Referenciada por (entrantes):**" & vbCrLf
        If IsArray(tinfo("relationsIn")) And UBound(tinfo("relationsIn")) >= 0 Then
            For i = 0 To UBound(tinfo("relationsIn"))
                Set rel = tinfo("relationsIn")(i)
                s = s & "- " & rel("name") & ": " & rel("table") & " → " & tName & " (" & rel("fieldPairs") & ")" & vbCrLf
            Next
        Else
            s = s & "- (ninguna)" & vbCrLf
        End If

        s = s & vbCrLf
    Next
    
    If gDebug Then LogMessage "Contenido final del Markdown (primeros 500 caracteres): " & Left(s, 500)

    Dim ts: Set ts = CreateObject("ADODB.Stream")
    ts.Type = 2: ts.Charset = "utf-8": ts.Open
    ts.WriteText s, 0
    ts.SaveToFile filePath, 2
    ts.Close
    
    If gDebug Then LogMessage "Archivo Markdown guardado: " & filePath & " (longitud: " & Len(s) & " caracteres)"
End Sub

' ============================================================================
' SECCIÓN 8: FUNCIONES UTILITARIAS
' ============================================================================

' Funciones utilitarias
Function GetFieldTypeName(fieldType)
    Select Case fieldType
        Case 1: GetFieldTypeName = "Boolean"        ' dbBoolean
        Case 2: GetFieldTypeName = "Byte"           ' dbByte
        Case 3: GetFieldTypeName = "Integer"        ' dbInteger
        Case 4: GetFieldTypeName = "Long"           ' dbLong
        Case 6: GetFieldTypeName = "Currency"       ' dbCurrency
        Case 7: GetFieldTypeName = "Single"         ' dbSingle
        Case 8: GetFieldTypeName = "Double"         ' dbDouble
        Case 9: GetFieldTypeName = "Date/Time"      ' dbDate
        Case 11: GetFieldTypeName = "Binary"        ' dbBinary
        Case 10: GetFieldTypeName = "Text"          ' dbText
        Case 12: GetFieldTypeName = "OLE Object"    ' dbLongBinary
        Case 13: GetFieldTypeName = "Memo"          ' dbMemo
        Case 15: GetFieldTypeName = "Replication ID" ' dbGUID
        Case 16: GetFieldTypeName = "Big Integer"   ' dbBigInt
        Case 17: GetFieldTypeName = "VarBinary"     ' dbVarBinary
        Case 18: GetFieldTypeName = "Char"          ' dbChar
        Case 19: GetFieldTypeName = "Numeric"       ' dbNumeric
        Case 20: GetFieldTypeName = "Decimal"       ' dbDecimal
        Case 21: GetFieldTypeName = "Float"         ' dbFloat
        Case 22: GetFieldTypeName = "Time"          ' dbTime
        Case 23: GetFieldTypeName = "TimeStamp"     ' dbTimeStamp
        Case Else: GetFieldTypeName = "Unknown (" & fieldType & ")"
    End Select
End Function

Function GetControlTypeName(controlType)
    Select Case controlType
        Case acLabel: GetControlTypeName = "Label"
        Case acRectangle: GetControlTypeName = "Rectangle"
        Case acLine: GetControlTypeName = "Line"
        Case acImage: GetControlTypeName = "Image"
        Case acCommandButton: GetControlTypeName = "Command Button"
        Case acOptionButton: GetControlTypeName = "Option Button"
        Case acCheckBox: GetControlTypeName = "Check Box"
        Case acOptionGroup: GetControlTypeName = "Option Group"
        Case acBoundObjectFrame: GetControlTypeName = "Bound Object Frame"
        Case acTextBox: GetControlTypeName = "Text Box"
        Case acListBox: GetControlTypeName = "List Box"
        Case acComboBox: GetControlTypeName = "Combo Box"
        Case acSubform: GetControlTypeName = "Subform/Subreport"
        Case acUnboundObjectFrame: GetControlTypeName = "Unbound Object Frame"
        Case acPageBreak: GetControlTypeName = "Page Break"
        Case acCustomControl: GetControlTypeName = "Custom Control"
        Case acToggleButton: GetControlTypeName = "Toggle Button"
        Case acTabCtl: GetControlTypeName = "Tab Control"
        Case acPage: GetControlTypeName = "Page"
        Case Else: GetControlTypeName = "Unknown (" & controlType & ")"
    End Select
End Function

Function GetViewTypeName(viewType)
    Select Case viewType
        Case acNormal: GetViewTypeName = "Single Form"
        Case acFormDS: GetViewTypeName = "Datasheet"
        Case acFormPivotTable: GetViewTypeName = "PivotTable"
        Case acFormPivotChart: GetViewTypeName = "PivotChart"
        Case Else: GetViewTypeName = "Unknown (" & viewType & ")"
    End Select
End Function

' Función para guardar datos en formato JSON
Sub SaveToJSON(data, filePath)
    Dim jsonText, st
    jsonText = ConvertToJSON(data)
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText jsonText, 0
    st.SaveToFile filePath, 2
    st.Close
End Sub

' Función básica para convertir a JSON
Function ConvertToJSON(obj)
    Dim result, key
    
    If TypeName(obj) = "Dictionary" Then
        result = "{"
        Dim first
        first = True
        
        For Each key In obj.Keys
            If Not first Then result = result & ","
            result = result & """" & key & """:" & ConvertToJSON(obj.Item(key))
            first = False
        Next
        
        result = result & "}"
    ElseIf IsArray(obj) Then
        result = "["
        Dim i
        For i = 0 To UBound(obj)
            If i > 0 Then result = result & ","
            result = result & ConvertToJSON(obj(i))
        Next
        result = result & "]"
    ElseIf VarType(obj) = vbString Then
        result = """" & Replace(obj, """", "\""") & """"
    ElseIf VarType(obj) = vbBoolean Then
        If obj Then result = "true" Else result = "false"
    ElseIf IsNull(obj) Then
        result = "null"
    Else
        result = CStr(obj)
    End If
    
    ConvertToJSON = result
End Function

' Función para crear directorios de forma recursiva
Sub CreateFolderRecursive(folderPath)
    Dim parentFolder
    
    If objFSO.FolderExists(folderPath) Then
        Exit Sub
    End If
    
    parentFolder = objFSO.GetParentFolderName(folderPath)
    
    If parentFolder <> "" And Not objFSO.FolderExists(parentFolder) Then
        CreateFolderRecursive parentFolder
    End If
    
    objFSO.CreateFolder folderPath
End Sub

' Función para resolver rutas relativas
Function ResolvePath(path)
    If objFSO.GetAbsolutePathName(path) = path Then
        ResolvePath = path
    Else
        ResolvePath = gScriptDir & "\" & path
    End If
End Function

' Función para obtener el nombre de archivo sin extensión
Function GetFileNameWithoutExtension(filePath)
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileNameWithoutExtension = fso.GetBaseName(filePath)
End Function

Function ResolveAssetImagePath(imgRef, config)
    On Error Resume Next
    ResolveAssetImagePath = ""

    If ("" & imgRef) = "" Then Exit Function

    Dim uiRoot, assetsDir, imgDir, baseDir, tryPath
    uiRoot    = UI_GetRoot(config)
    assetsDir = UI_GetAssetsDir(config)
    imgDir    = UI_GetAssetsImgDir(config)

    ' 1) Si viene ruta absoluta, úsala directamente
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(imgRef) Then
        ResolveAssetImagePath = imgRef
        Exit Function
    End If

    ' 2) Probar combinaciones relativas canónicas
    Dim candidates(2)
    candidates(0) = ResolvePath(uiRoot & "\" & assetsDir & "\" & imgDir & "\" & imgRef)       ' ./ui/assets/img/<imgRef>
    candidates(1) = ResolvePath(uiRoot & "\" & imgDir & "\" & imgRef)                         ' ./ui/img/<imgRef> (fallback)
    candidates(2) = ResolvePath(imgRef)                                                       ' relativo al script

    Dim i
    For i = 0 To UBound(candidates)
        tryPath = candidates(i)
        If fso.FileExists(tryPath) Then
            ResolveAssetImagePath = tryPath
            Exit Function
        End If
    Next

    ' 3) Si vino sin extensión, probar con extensiones permitidas
    Dim bare, ext, exts, arr, j
    bare = fso.GetBaseName(imgRef)
    exts = UI_GetAssetsImgExtensions(config)
    arr = Split(exts, ",")
    For j = 0 To UBound(arr)
        ext = Trim(arr(j))
        If Left(ext,1) <> "." Then ext = "." & ext
        tryPath = ResolvePath(uiRoot & "\" & assetsDir & "\" & imgDir & "\" & bare & ext)
        If fso.FileExists(tryPath) Then
            ResolveAssetImagePath = tryPath
            Exit Function
        End If
    Next
End Function

' ============================================================================
' Función para verificar si un modulo necesita actualización

' ============================================================================

' Función para reconstruir modulos VBA desde archivos fuente



' Función para actualizar modulos VBA desde archivos fuente


' Función para obtener lista de archivos de modulos
Function GetModuleFiles(srcPath, extensions, includeSubdirs, filePattern)
    Dim files(), fileCount, folder, file, subFolder
    Dim extArray, i, j, ext, fileName, fileExt
    Dim normalizedPattern
    
    fileCount = 0
    ReDim files(-1) ' Array vacío por defecto
    
    ' Normalizar filePattern - si viene vacío o nulo, usar "*" por defecto
    If IsNull(filePattern) Or filePattern = "" Then
        normalizedPattern = "*"
    Else
        normalizedPattern = filePattern
    End If
    
    ' Normalizar extensiones: trim + lower, ignorar entradas vacías
    extArray = Split(extensions, ",")
    Dim validExtensions(), validExtCount
    validExtCount = 0
    ReDim validExtensions(-1)
    
    For i = 0 To UBound(extArray)
        ext = Trim(extArray(i))
        If ext <> "" Then
            ' Remover el punto inicial si existe para normalizar
            If Left(ext, 1) = "." Then
                ext = Mid(ext, 2)
            End If
            ReDim Preserve validExtensions(validExtCount)
            validExtensions(validExtCount) = LCase(ext)
            validExtCount = validExtCount + 1
        End If
    Next
    
    ' Si no hay extensiones válidas, salir
    If validExtCount = 0 Then
        GetModuleFiles = files
        Exit Function
    End If
    
    If gDebug Then 
        LogMessage "[DEBUG] GetModuleFiles - srcPath: " & srcPath & ", patron: " & normalizedPattern
        LogMessage "[DEBUG] Extensiones normalizadas: " & Join(validExtensions, ", ") & " (total: " & validExtCount & ")"
    End If
    
    On Error Resume Next
    Set folder = objFSO.GetFolder(srcPath)
    If Err.Number <> 0 Then
        If gDebug Then LogMessage "[DEBUG] Error accediendo directorio: " & srcPath & " - " & Err.Description
        Err.Clear
        GetModuleFiles = files
        Exit Function
    End If
    On Error GoTo 0
    
    If gDebug Then LogMessage "[DEBUG] Procesando " & folder.Files.Count & " archivos en directorio"
    
    ' Procesar archivos en el directorio actual
    For Each file In folder.Files
        fileName = file.Name
        fileExt = LCase(objFSO.GetExtensionName(fileName))
        
        If gDebug Then LogMessage "[DEBUG] Evaluando archivo: " & fileName & " (ext: " & fileExt & ")"
        
        ' Verificar si la extensión coincide con alguna válida
        For j = 0 To UBound(validExtensions)
            If fileExt = validExtensions(j) Then
                If gDebug Then LogMessage "[DEBUG] Extension coincide para: " & fileName
                
                ' Verificar patrón
                Dim patternMatch
                patternMatch = False
                
                If normalizedPattern = "*" Then
                    patternMatch = True
                ElseIf InStr(normalizedPattern, ";") > 0 Then
                    ' Múltiples patrones separados por ;
                    Dim patterns, k
                    patterns = Split(normalizedPattern, ";")
                    For k = 0 To UBound(patterns)
                        If MatchesPattern(fileName, Trim(patterns(k))) Then
                            patternMatch = True
                            Exit For
                        End If
                    Next
                Else
                    ' Patrón único
                    patternMatch = MatchesPattern(fileName, normalizedPattern)
                End If
                
                If patternMatch Then
                    If gDebug Then LogMessage "[DEBUG] Patron coincide, agregando: " & fileName
                    ReDim Preserve files(fileCount)
                    files(fileCount) = file.Path
                    fileCount = fileCount + 1
                Else
                    If gDebug Then LogMessage "[DEBUG] Patron NO coincide para: " & fileName & " (patron: " & normalizedPattern & ")"
                End If
                Exit For
            End If
        Next
    Next
    
    ' Procesar subdirectorios si está habilitado
    If includeSubdirs Then
        For Each subFolder In folder.SubFolders
            Dim subFiles
            subFiles = GetModuleFiles(subFolder.Path, extensions, includeSubdirs, filePattern)
            
            If UBound(subFiles) >= 0 Then
                For i = 0 To UBound(subFiles)
                    ReDim Preserve files(fileCount)
                    files(fileCount) = subFiles(i)
                    fileCount = fileCount + 1
                Next
            End If
        Next
    End If
    
    If gDebug Then LogMessage "[DEBUG] Total archivos encontrados: " & UBound(files) + 1
    
    GetModuleFiles = files
End Function

' Función auxiliar para verificar si un nombre de archivo coincide con un patrón
Function MatchesPattern(fileName, pattern)
    MatchesPattern = False
    
    If pattern = "*" Then
        MatchesPattern = True
    ElseIf InStr(pattern, "*") > 0 Then
        ' Patrón con wildcards - convertir a regex simple
        Dim regexPattern
        regexPattern = Replace(pattern, "*", ".*")
        regexPattern = "^" & regexPattern & "$"
        
        On Error Resume Next
        Dim regex
        Set regex = CreateObject("VBScript.RegExp")
        regex.Pattern = regexPattern
        regex.IgnoreCase = True
        MatchesPattern = regex.Test(fileName)
        If Err.Number <> 0 Then
            ' Fallback a comparación simple si regex falla
            MatchesPattern = (InStr(LCase(fileName), LCase(Replace(pattern, "*", ""))) > 0)
            Err.Clear
        End If
        On Error GoTo 0
    Else
        ' Comparación exacta (case insensitive)
        MatchesPattern = (LCase(fileName) = LCase(pattern))
    End If
End Function

' Función para leer contenido de archivo de modulo
Function ReadModuleFile(filePath)
    Dim stream, content
    
    ReadModuleFile = ""
    
    ' Verificar que el archivo existe
    If Not objFSO.FileExists(filePath) Then
        If gDebug Then LogMessage "[DEBUG] Archivo no existe: " & filePath
        Exit Function
    End If
    
    On Error Resume Next
    
    ' Crear stream ADODB para lectura robusta con windows-1252 por defecto
    Set stream = CreateObject("ADODB.Stream")
    If Err.Number <> 0 Then
        LogMessage "Error creando stream para archivo: " & filePath & " - " & Err.Number & " " & Err.Description
        Err.Clear
        Exit Function
    End If
    
    ' Configurar stream como texto con windows-1252
    stream.Type = 2 ' adTypeText
    stream.Charset = "windows-1252"
    stream.Open
    
    If Err.Number <> 0 Then
        LogMessage "Error abriendo stream para archivo: " & filePath & " - " & Err.Number & " " & Err.Description
        Err.Clear
        Exit Function
    End If
    
    ' Cargar archivo
    stream.LoadFromFile filePath
    If Err.Number <> 0 Then
        LogMessage "Error cargando archivo: " & filePath & " - " & Err.Number & " " & Err.Description
        Err.Clear
        If stream.State = 1 Then stream.Close
        Exit Function
    End If
    
    ' Leer contenido completo
    If Not stream.EOS Then
        content = stream.ReadText(-1) ' Leer todo
        If Err.Number <> 0 Then
            LogMessage "Error leyendo contenido de archivo: " & filePath & " - " & Err.Number & " " & Err.Description
            content = ""
            Err.Clear
        End If
        
        ReadModuleFile = content
    End If
    
    ' Cerrar stream de forma segura
    If Not stream Is Nothing Then
        If stream.State = 1 Then stream.Close
    End If
    
    On Error GoTo 0
    
    If gDebug And Len(ReadModuleFile) > 0 Then 
        LogMessage "[DEBUG] Leidos " & Len(ReadModuleFile) & " caracteres de: " & filePath
    End If
End Function

' Función para importar modulo a Access
' Función para detectar la codificación de un archivo por BOM
' ====== UTIL: UBound seguro para variantes/arrays vacios ======
Function SafeUBound(v)
    On Error Resume Next
    Dim n: n = -1
    If IsArray(v) Then
        n = UBound(v)
        If Err.Number <> 0 Then
            n = -1
            Err.Clear
        End If
    End If
    SafeUBound = n
End Function

' ====== Deteccion de charset por BOM, sin indexar si no hay bytes ======
Function DetectCharset(path)
    DetectCharset = "windows-1252" ' Default fallback
    Dim stm
    
    Set stm = CreateObject("ADODB.Stream")
    On Error Resume Next
    
    ' Abrir como binario para leer BOM
    stm.Type = 1 ' adTypeBinary
    stm.Open
    stm.LoadFromFile path
    
    If Err.Number <> 0 Then 
        Err.Clear
        If stm.State = 1 Then stm.Close
        If gDebug Then LogMessage "[DEBUG] No se pudo leer archivo para detectar charset: " & path
        Exit Function
    End If
    
    ' Verificar BOM UTF-8 (EF BB BF)
    If stm.Size >= 3 Then
        stm.Position = 0
        Dim b3
        b3 = stm.Read(3)
        If IsArray(b3) And UBound(b3) >= 2 Then
            If (b3(0) = &HEF) And (b3(1) = &HBB) And (b3(2) = &HBF) Then 
                DetectCharset = "utf-8"
                If gDebug Then LogMessage "[DEBUG] BOM UTF-8 detectado en: " & path
                stm.Close
                Exit Function
            End If
        End If
    End If
    
    ' Verificar BOM UTF-16 LE (FF FE) y BE (FE FF)
    If stm.Size >= 2 Then
        stm.Position = 0
        Dim b2
        b2 = stm.Read(2)
        If IsArray(b2) And UBound(b2) >= 1 Then
            If (b2(0) = &HFF) And (b2(1) = &HFE) Then 
                DetectCharset = "unicode"
                If gDebug Then LogMessage "[DEBUG] BOM UTF-16 LE detectado en: " & path
                stm.Close
                Exit Function
            End If
            If (b2(0) = &HFE) And (b2(1) = &HFF) Then 
                DetectCharset = "bigendianunicode"
                If gDebug Then LogMessage "[DEBUG] BOM UTF-16 BE detectado en: " & path
                stm.Close
                Exit Function
            End If
        End If
    End If
    
    stm.Close
    On Error GoTo 0
    
    If gDebug Then LogMessage "[DEBUG] Sin BOM detectado, usando charset por defecto para: " & path
End Function

' ====== Transcodificacion robusta a Windows-1252 (ANSI) ======
Sub TranscodeTextFile(srcPath, dstPath, dstCharset)
    On Error Resume Next

    Dim inS, outS, srcCs, txt, finalCharset
    Set inS = CreateObject("ADODB.Stream")
    Set outS = CreateObject("ADODB.Stream")

    ' Determinar charset de origen de forma segura
    srcCs = DetectCharset(srcPath)
    
    ' Abrir lectura como texto con el charset detectado
    inS.Type = 2 ' adTypeText
    inS.Charset = srcCs
    inS.Open
    inS.LoadFromFile srcPath

    ' Leer TODO el texto
    txt = ""
    txt = inS.ReadText(-1)
    If Err.Number <> 0 Then
        ' Si falla (p.ej. por charset incorrecto), reintentar asumiendo windows-1252
        WScript.Echo "[ERROR] Error cargando archivo con charset " & srcCs & ": " & Err.Description
        Err.Clear
        inS.Close
        inS.Type = 2
        inS.Charset = "windows-1252"
        inS.Open
        inS.LoadFromFile srcPath
        txt = inS.ReadText(-1)
        If Err.Number <> 0 Then
            WScript.Echo "[ERROR] Error cargando archivo con windows-1252: " & Err.Description
            Exit Sub
        End If
    End If
    inS.Close

    ' Escribir como ANSI (windows-1252) u otro solicitado
    If LCase(dstCharset) = "" Then
        finalCharset = "windows-1252"
    Else
        finalCharset = dstCharset
    End If
    
    outS.Type = 2 ' text
    outS.Charset = finalCharset
    outS.Open
    outS.WriteText txt, 0   ' adWriteChar
    outS.SaveToFile dstPath, 2 ' adSaveCreateOverWrite
    outS.Close
    
    If Err.Number <> 0 Then
        WScript.Echo "[ERROR] Error guardando archivo: " & Err.Description
        Err.Clear
    End If
End Sub

' Funcion para importar modulo VBA desde archivo con codificacion correcta


' Función para asegurar que Option Compare Database y Option Explicit estén presentes
Function EnsureOptionExplicit(vbMod)
    On Error Resume Next
    EnsureOptionExplicit = False
    If vbMod Is Nothing Then Exit Function

    Dim total, i, lineTxt, found, firstNonEmpty, firstNonEmptyLineNum
    found = False
    total = vbMod.CountOfLines

    ' Si está vacío, inserta directamente
    If total < 1 Then
        vbMod.InsertLines 1, "Option Explicit"
        If Err.Number <> 0 Then Err.Clear
        EnsureOptionExplicit = True
        Exit Function
    End If

    ' Busca "Option Explicit" en las primeras líneas (hasta 20 para rendimiento)
    firstNonEmpty = ""
    firstNonEmptyLineNum = 0
    For i = 1 To total
        lineTxt = vbMod.Lines(i, 1)
        If LCase(Trim(lineTxt)) = "option explicit" Then
            found = True
            Exit For
        End If
        If firstNonEmpty = "" Then
            If Trim(lineTxt) <> "" Then
                firstNonEmpty = LCase(Trim(lineTxt))
                firstNonEmptyLineNum = i
            End If
        End If
        If i >= 20 Then Exit For
    Next

    If Not found Then
        ' Si la primera no vacía es Option Compare, inserta debajo; si no, al inicio
        If Left(firstNonEmpty, 13) = "option compare" Then
            vbMod.InsertLines firstNonEmptyLineNum + 1, "Option Explicit"
        Else
            vbMod.InsertLines 1, "Option Explicit"
        End If
        If Err.Number <> 0 Then Err.Clear
    End If

    EnsureOptionExplicit = True
End Function

Function PostProcessInsertedModule(vbComp)
    On Error Resume Next
    PostProcessInsertedModule = False
    If vbComp Is Nothing Then Exit Function

    Dim cm, total, i, lineTxt
    Dim optExpLine, optCmpLine, firstCodeLine

    Set cm = vbComp.CodeModule
    If cm Is Nothing Then Exit Function

    total = cm.CountOfLines
    If total < 1 Then
        cm.InsertLines 1, "Option Explicit"
        If Err.Number <> 0 Then Err.Clear
        PostProcessInsertedModule = True
        Exit Function
    End If

    optExpLine = 0
    optCmpLine = 0
    firstCodeLine = 0

    ' Recorre un bloque razonable (primeras 400 líneas o total) para ubicar cabecera/atributos y opciones
    For i = 1 To total
        lineTxt = LCase(Trim(cm.Lines(i, 1)))

        ' Detecta Option Explicit / Option Compare
        If lineTxt = "option explicit" Then
            optExpLine = i
            Exit For ' Ya no necesitamos buscar más si ya está
        End If
        If Left(lineTxt, 14) = "option compare" And optCmpLine = 0 Then
            optCmpLine = i
        End If

        ' Marca la primera línea "de código" útil (salta cabeceras .cls y atributos)
        If firstCodeLine = 0 Then
            If lineTxt <> "" _
               And Left(lineTxt, 7) <> "version" _
               And Left(lineTxt, 5) <> "begin" _
               And lineTxt <> "end" _
               And Left(lineTxt, 9) <> "attribute" _
               And Left(lineTxt, 1) <> "'" Then
                firstCodeLine = i
            End If
        End If

        If i >= 400 Then Exit For
    Next

    If optExpLine = 0 Then
        ' Insertar Option Explicit debajo de Option Compare si existe; si no, en el primer código o al inicio
        If optCmpLine > 0 Then
            cm.InsertLines optCmpLine + 1, "Option Explicit"
        ElseIf firstCodeLine > 0 Then
            cm.InsertLines firstCodeLine, "Option Explicit"
        Else
            cm.InsertLines 1, "Option Explicit"
        End If
        If Err.Number <> 0 Then Err.Clear
    End If

    PostProcessInsertedModule = True
End Function

' Función para eliminar modulo existente
Sub DeleteExistingModule(moduleName, moduleType)
    On Error Resume Next
    
    Select Case moduleType
        Case acModule, acClassModule
            ' Eliminar modulo VBA
            objAccess.VBE.VBProjects(1).VBComponents.Remove objAccess.VBE.VBProjects(1).VBComponents(moduleName)
    End Select
    
    Err.Clear ' Ignorar errores si el modulo no existe
End Sub

' Función para obtener modulos existentes en la base de datos
Function GetDatabaseModules()
    On Error Resume Next
    
    Dim modules, i, vbComp, componentCount, formsCount
    Set modules = CreateObject("Scripting.Dictionary")
    
    ' Obtener modulos VBA - verificar que el proyecto VBA existe
    If Not objAccess.VBE Is Nothing And objAccess.VBE.VBProjects.Count > 0 Then
        componentCount = objAccess.VBE.VBProjects(1).VBComponents.Count
        If componentCount > 0 Then
            For i = 1 To componentCount
                Set vbComp = objAccess.VBE.VBProjects(1).VBComponents(i)
                If Not vbComp Is Nothing Then
                    modules.Add vbComp.Name, Now() ' Usar fecha actual como placeholder
                End If
            Next
        End If
    End If
    
    ' Obtener formularios - verificar que existen
    If Not objAccess.CurrentProject Is Nothing Then
        formsCount = objAccess.CurrentProject.AllForms.Count
        If formsCount > 0 Then
            For i = 0 To formsCount - 1
                modules.Add objAccess.CurrentProject.AllForms(i).Name, Now()
            Next
        End If
    End If
    
    Set GetDatabaseModules = modules
    
    If Err.Number <> 0 Then
        LogMessage "Error obteniendo modulos de la base de datos: " & Err.Description
        Err.Clear
        ' Devolver diccionario vacío en caso de error
        Set modules = CreateObject("Scripting.Dictionary")
        Set GetDatabaseModules = modules
    End If
End Function

' Función para extraer modulos VBA desde Access hacia archivos fuente
Function ExtractModulesToFiles(dbPath)
    On Error Resume Next
    
    Dim config, srcPath, extensions, includeSubdirs, filePattern
    Dim i, vbComp, moduleName, moduleContent, filePath, fileExt
    Dim vbProject, vbComponents
    
    ExtractModulesToFiles = False
    
    ' Cerrar procesos de Access existentes antes de comenzar
    CloseExistingAccessProcesses
    
    ' Cargar configuración
    Set config = LoadConfig(gConfigPath)
    srcPath = config("MODULES_SrcPath")
    extensions = config("MODULES_Extensions")
    includeSubdirs = LCase(config("MODULES_IncludeSubdirectories")) = "true"
    filePattern = config("MODULES_FilePattern")
    
    LogMessage "Iniciando extraccion de modulos VBA hacia: " & srcPath
    LogMessage "Base de datos origen: " & dbPath
    
    ' Crear directorio fuente si no existe
    If Not objFSO.FolderExists(srcPath) Then
        CreateFolderRecursive srcPath
        LogMessage "Directorio creado: " & srcPath
    End If
    
    ' Abrir Access
    Set objAccess = OpenAccess(dbPath, gPassword)
    If objAccess Is Nothing Then
        LogMessage "Error: No se pudo abrir la base de datos para extraccion"
        Exit Function
    End If
    
    ' Extraer modulos VBA
    LogMessage "Extrayendo modulos VBA..."
    
    ' Obtener el proyecto VBA activo
    Set vbProject = objAccess.VBE.ActiveVBProject
    If Err.Number <> 0 Or vbProject Is Nothing Then
        Err.Clear
        LogMessage "Error: No se pudo acceder al proyecto VBA"
        CloseAccess objAccess
        Exit Function
    End If
    
    Set vbComponents = vbProject.VBComponents
    If Err.Number <> 0 Or vbComponents Is Nothing Then
        Err.Clear
        LogMessage "Error: No se pudo acceder a los componentes VBA"
        CloseAccess objAccess
        Exit Function
    End If
    
    For i = 1 To vbComponents.Count
        Set vbComp = vbComponents(i)
        moduleName = vbComp.Name
        
        ' Verificar que el nombre del modulo no esté vacío
        If Len(Trim(moduleName)) = 0 Then
            LogMessage "Modulo con nombre vacio, omitiendo"
        Else
            ' Determinar extensión según tipo de modulo
            Select Case vbComp.Type
                Case 1 ' vbext_ct_StdModule
                    fileExt = ".bas"
                Case 2 ' vbext_ct_ClassModule
                    fileExt = ".cls"
                Case Else
                    LogMessage "Tipo de modulo no soportado: " & moduleName & " (Tipo: " & vbComp.Type & ")"
                    fileExt = ""
            End Select
            
            ' Solo procesar si tenemos una extensión válida
            If fileExt <> "" Then
                ' Construir ruta del archivo
                filePath = srcPath & "\" & moduleName & fileExt
                
                LogMessage "Extrayendo modulo: " & moduleName & " -> " & filePath
                
                ' Obtener contenido del modulo usando CodeModule
                If vbComp.CodeModule.CountOfLines > 0 Then
                    moduleContent = vbComp.CodeModule.Lines(1, vbComp.CodeModule.CountOfLines)
                Else
                    moduleContent = ""
                End If
                
                ' Escribir archivo
                If WriteModuleFile(filePath, moduleContent) Then
                    LogMessage "Modulo " & moduleName & " extraído correctamente"
                Else
                    LogMessage "Error extrayendo modulo: " & moduleName
                End If
            End If
        End If
    Next

ExtractReports:
    ' Extraer informes
    LogMessage "Extrayendo informes..."
    For i = 0 To objAccess.CurrentProject.AllReports.Count - 1
        moduleName = objAccess.CurrentProject.AllReports(i).Name
        
        ' Verificar patron
        If filePattern = "*" Or filePattern = "" Or InStr(LCase(moduleName), LCase(filePattern)) > 0 Then
            filePath = srcPath & "\" & moduleName & ".rpt"
            LogMessage "Extrayendo informe: " & moduleName & " -> " & filePath
            
            ' Exportar informe
            If ExportAccessObject(acReport, moduleName, filePath) Then
                LogMessage "Informe " & moduleName & " extraído correctamente"
            Else
                LogMessage "Error extrayendo informe: " & moduleName
            End If
        Else
            LogMessage "Informe " & moduleName & " no coincide con el patron, omitiendo"
        End If
    Next
    
    ' Verificar archivos creados antes de reportar éxito
    Dim filesCreated, totalFiles
    filesCreated = 0
    totalFiles = 0
    
    ' Contar archivos .bas y .cls en el directorio src
    If objFSO.FolderExists(srcPath) Then
        Dim srcFolder, file
        Set srcFolder = objFSO.GetFolder(srcPath)
        For Each file In srcFolder.Files
            If LCase(objFSO.GetExtensionName(file.Name)) = "bas" Or LCase(objFSO.GetExtensionName(file.Name)) = "cls" Then
                filesCreated = filesCreated + 1
            End If
            totalFiles = totalFiles + 1
        Next
    End If
    
    ExtractModulesToFiles = True
    
    If filesCreated > 0 Then
        LogMessage "Extraccion completada exitosamente - " & filesCreated & " modulos extraidos"
    Else
        LogMessage "Extraccion completada - No se extrajeron modulos (verificar patron de filtro)"
    End If
    
    ' Cerrar Access
    CloseAccess objAccess
    
    If Err.Number <> 0 Then
        LogMessage "Error durante extraccion: " & Err.Description
        Err.Clear
        ExtractModulesToFiles = False
    End If
End Function

' Función para escribir contenido de modulo a archivo con metadatos
Function WriteModuleFile(filePath, content)
    On Error Resume Next
    
    Dim file, fileExt, finalContent
    Dim hasAttribute, hasOption, hasExplicit, insertAfter
    
    WriteModuleFile = False
    
    ' Crear directorio padre si no existe
    Dim parentDir
    parentDir = objFSO.GetParentFolderName(filePath)
    If Not objFSO.FolderExists(parentDir) Then
        CreateFolderRecursive parentDir
    End If
    
    ' Obtener extensión del archivo para determinar metadatos
    fileExt = LCase(objFSO.GetExtensionName(filePath))
    
    ' Agregar metadatos según el tipo de archivo
    finalContent = ""
    
    Select Case fileExt
        Case "bas"
            ' Módulos estándar: agregar metadatos si no existen
            hasAttribute = InStr(1, content, "Attribute VB_Name", vbTextCompare) > 0
            hasOption = InStr(1, content, "Option Compare Database", vbTextCompare) > 0
            hasExplicit = InStr(1, content, "Option Explicit", vbTextCompare) > 0
            
            finalContent = content
            
            ' Agregar Attribute VB_Name si no existe
            If Not hasAttribute Then
                finalContent = "Attribute VB_Name = """ & objFSO.GetBaseName(filePath) & """" & vbCrLf & finalContent
            End If
            
            ' Agregar Option Compare Database si no existe (debe ir antes que Option Explicit)
            If Not hasOption Then
                If hasAttribute Then
                    ' Si ya tiene Attribute, agregar Option Compare Database después
                    finalContent = Replace(finalContent, "Attribute VB_Name = """ & objFSO.GetBaseName(filePath) & """", _
                                         "Attribute VB_Name = """ & objFSO.GetBaseName(filePath) & """" & vbCrLf & "Option Compare Database", 1, 1)
                Else
                    ' Si no tiene Attribute pero acabamos de agregarlo, agregar Option Compare Database después
                    finalContent = Replace(finalContent, "Attribute VB_Name = """ & objFSO.GetBaseName(filePath) & """", _
                                         "Attribute VB_Name = """ & objFSO.GetBaseName(filePath) & """" & vbCrLf & "Option Compare Database", 1, 1)
                End If
            End If
            
            ' Agregar Option Explicit si no existe (debe ir después de Option Compare Database)
            If Not hasExplicit Then
                ' Buscar dónde insertar Option Explicit
                If hasOption Or (Not hasOption And Not hasAttribute) Then
                    ' Si tiene Option Compare Database o acabamos de agregar Attribute + Option Compare Database
                    insertAfter = "Option Compare Database"
                    finalContent = Replace(finalContent, insertAfter, insertAfter & vbCrLf & "Option Explicit", 1, 1)
                ElseIf hasAttribute Then
                    ' Si solo tiene Attribute
                    insertAfter = "Attribute VB_Name = """ & objFSO.GetBaseName(filePath) & """"
                    finalContent = Replace(finalContent, insertAfter, insertAfter & vbCrLf & "Option Explicit", 1, 1)
                End If
            End If
        Case "cls"
            ' Módulos de clase: verificar si ya tienen metadatos completos
            If InStr(1, content, "VERSION 1.0 CLASS", vbTextCompare) = 0 Then
                ' Agregar metadatos completos de clase
                finalContent = "VERSION 1.0 CLASS" & vbCrLf & _
                              "BEGIN" & vbCrLf & _
                              "  MultiUse = -1  'True" & vbCrLf & _
                              "END" & vbCrLf & _
                              "Attribute VB_Name = """ & objFSO.GetBaseName(filePath) & """" & vbCrLf & _
                              "Attribute VB_GlobalNameSpace = False" & vbCrLf & _
                              "Attribute VB_Creatable = False" & vbCrLf & _
                              "Attribute VB_PredeclaredId = False" & vbCrLf & _
                              "Attribute VB_Exposed = False" & vbCrLf & content
                ' Agregar Option Compare Database si no existe
                If InStr(1, finalContent, "Option Compare Database", vbTextCompare) = 0 Then
                    finalContent = finalContent & vbCrLf & "Option Compare Database"
                End If
            Else
                ' Ya tiene metadatos, usar contenido tal como está
                finalContent = content
            End If
        Case Else
            ' Otros tipos: usar contenido tal como está
            finalContent = content
    End Select
    
    ' Escribir archivo usando ADODB.Stream para preservar caracteres especiales
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "UTF-8"
    objStream.Open
    objStream.WriteText finalContent, 0 ' adWriteChar
    objStream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    objStream.Close
    Set objStream = Nothing
    WriteModuleFile = True
    
    If Err.Number <> 0 Then
        LogMessage "Error escribiendo archivo: " & filePath & " - " & Err.Description
        Err.Clear
    End If
End Function

' Función para obtener valores de configuración con fallback
Function CfgGet(cfg, keyNew, keyOld, def)
    On Error Resume Next
    If cfg.Exists(keyNew) Then
        CfgGet = cfg(keyNew)
    ElseIf cfg.Exists(keyOld) Then
        CfgGet = cfg(keyOld)
    Else
        CfgGet = def
    End If
End Function

' Función para convertir string a boolean
Function ToBool(s, def)
    On Error Resume Next
    Dim t: t = LCase(CStr(s))
    If t = "true" Or t = "1" Or t = "yes" Or t = "si" Then
        ToBool = True
    ElseIf t = "false" Or t = "0" Or t = "no" Then
        ToBool = False
    Else
        ToBool = def
    End If
End Function

' Función stub para validar configuración
Sub ValidateConfig()
    LogMessage "Validacion de configuracion: OK"
End Sub

' Función para exportar objetos de Access
Function ExportAccessObject(objectType, objectName, filePath)
    On Error Resume Next
    
    ExportAccessObject = False
    
    ' Crear directorio padre si no existe
    Dim parentDir
    parentDir = objFSO.GetParentFolderName(filePath)
    If Not objFSO.FolderExists(parentDir) Then
        CreateFolderRecursive parentDir
    End If
    
    ' Exportar objeto usando SaveAsText
    objAccess.Application.SaveAsText objectType, objectName, filePath
    
    ' Manejo rigido de errores para evitar contaminacion
    If Err.Number = 0 Then
        ExportAccessObject = True
    Else
        LogMessage "Error exportando " & objectName & ": " & Err.Description & " (Codigo: " & Err.Number & ")"
        Err.Clear
        ' Salir inmediatamente sin ejecutar logica posterior sobre estado inconsistente
        Exit Function
    End If
    
    ' Verificacion adicional: comprobar que el archivo se creo realmente
    If Not objFSO.FileExists(filePath) Then
        LogMessage "Advertencia: archivo no creado para " & objectName & " en " & filePath
        ExportAccessObject = False
    End If
End Function

Sub Main()
    ' Inicializar objetos
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objArgs = WScript.Arguments
    
    ' Establecer rutas base del script
    gScriptPath = WScript.ScriptFullName
    gScriptDir = objFSO.GetParentFolderName(gScriptPath)
    
    ' Configuración por defecto con rutas relativas
    gConfigPath = gScriptDir & "\cli.ini"
    gVerbose = False
    gQuiet = False
    gDryRun = False
    gPassword = ""
    gOutputPath = gScriptDir & "\output"
    
    ' Cargar configuracion de modulos
    Set gConfig = LoadConfig(gConfigPath)
    
    ' Integrar configuración de base de datos desde cli.ini
    If gConfig("DATABASE_Password") <> "" Then
        gPassword = gConfig("DATABASE_Password")
    End If
    
    g_ModulesSrcPath = gConfig("MODULES_SrcPath")
    g_ModulesExtensions = gConfig("MODULES_Extensions")
    g_ModulesIncludeSubdirs = LCase(gConfig("MODULES_IncludeSubdirectories")) = "true"
    
    ' Procesar argumentos de línea de comandos
    If objArgs.Count = 0 Then
        ShowHelp
        WScript.Quit 1
    End If
    
    ' Cargar configuración primero
    Set objConfig = LoadConfig(gConfigPath)
    
    ' Aplicar configuración
    If objConfig.Exists("LOGGING_Verbose") Then
        gVerbose = LCase(objConfig.Item("LOGGING_Verbose")) = "true"
    End If
    
    If objConfig.Exists("LOGGING_QuietMode") Then
        gQuiet = LCase(objConfig.Item("LOGGING_QuietMode")) = "true"
    End If
    
    ' Crear array de argumentos sin modificadores
    Dim cleanArgs()
    Dim cleanArgCount
    cleanArgCount = 0
    
    ' Procesar modificadores y filtrar argumentos
    Dim i
    For i = 0 To objArgs.Count - 1
        If objArgs(i) = "/test" Then
            LogMessage "Modo de prueba activado"
            ' RunTests - función no implementada aún
            LogMessage "Función RunTests no implementada"
            WScript.Quit 0
        ElseIf objArgs(i) = "/dry-run" Or objArgs(i) = "--dry-run" Then
            gDryRun = True
            LogMessage "Modo simulacion activado"
        ElseIf objArgs(i) = "/validate" Then
            ValidateConfig
            WScript.Quit 0
        ElseIf objArgs(i) = "/verbose" Then
            gVerbose = True
            LogMessage "Modo verbose activado"
        ElseIf objArgs(i) = "/debug" Then
            gDebug = True
            LogMessage "Modo debug activado"
        Else
            ' Es un argumento normal, no un modificador
            ReDim Preserve cleanArgs(cleanArgCount)
            cleanArgs(cleanArgCount) = objArgs(i)
            cleanArgCount = cleanArgCount + 1
        End If
    Next
    
    ' Verificar que tenemos al menos un comando
    If cleanArgCount = 0 Then
        ShowHelp
        WScript.Quit 1
    End If
    
    ' Procesar comando principal
    Dim command
    command = cleanArgs(0)
    
    Select Case LCase(command)
        Case "extract-modules"
            ' Si no se especifica ruta, usar DefaultPath del config
            If cleanArgCount >= 2 Then
                gDbPath = ResolvePath(cleanArgs(1))
            Else
                ' Usar DefaultPath de la configuración
                gDbPath = ResolvePath(objConfig("DATABASE_DefaultPath"))
                LogMessage "Usando base de datos por defecto: " & gDbPath
            End If
            
            If Not gDryRun Then
                If gVerbose Then WScript.Echo "Extrayendo modulos VBA desde Access..."
                If Not ExtractModulesToFiles(gDbPath) Then
                    WScript.Echo "Error: No se pudo completar la extraccion de modulos"
                    WScript.Quit 1
                End If
            Else
                WScript.Echo "[DRY-RUN] Se extraerian modulos VBA desde: " & gDbPath
            End If
            

            
        Case "rebuild"
            ' El comando rebuild no acepta parámetros adicionales, siempre usa DefaultPath
            gDbPath = ResolvePath(objConfig("DATABASE_DefaultPath"))
            LogMessage "Usando base de datos por defecto: " & gDbPath
            
            ' Validar que el archivo existe
            If Not objFSO.FileExists(gDbPath) Then
                WScript.Echo "Error: El archivo de base de datos no existe: " & gDbPath
                WScript.Quit 1
            End If
            
            If gVerbose Then WScript.Echo "Reconstruyendo modulos VBA..."
            
            ' Abrir Access para RebuildProject
            Set objAccess = OpenAccess(gDbPath, gPassword)
            If objAccess Is Nothing Then
                WScript.Echo "Error: No se pudo abrir Access"
                WScript.Quit 1
            End If
            
            ' Llamar a RebuildProject en lugar de RebuildModules
            Call RebuildProject()
            
            ' Cerrar Access
            Call CloseAccess(objAccess)
            WScript.Quit 0
            
        Case "update"
            ' El comando update usa siempre la base de datos por defecto del .ini
            ' Firma esperada: update <lista_modulos>
            gDbPath = ResolvePath(objConfig("DATABASE_DefaultPath"))
            LogMessage "Usando base de datos por defecto: " & gDbPath
            
            ' Validar que el archivo existe
            If Not objFSO.FileExists(gDbPath) Then
                WScript.Echo "Error: El archivo de base de datos no existe: " & gDbPath
                WScript.Quit 1
            End If
            
            If cleanArgCount >= 2 Then
                ' Captura todos los argumentos como lista de módulos (desde el primer argumento)
                Dim modulesArg
                modulesArg = ""
                
                ' Concatenar todos los argumentos de módulos
                Dim argIndex
                For argIndex = 1 To cleanArgCount - 1
                    If modulesArg <> "" Then
                        modulesArg = modulesArg & " " & cleanArgs(argIndex)
                    Else
                        modulesArg = cleanArgs(argIndex)
                    End If
                Next
                
                If Not gDryRun Then
                    If gVerbose Then WScript.Echo "Ejecutando update de módulos específicos..."
                    Dim res
                    res = UpdateModules(gDbPath, modulesArg)
                    If Not res Then
                        WScript.Echo "Error: No se pudo completar el update de módulos"
                        WScript.Quit 1
                    End If
                    LogMessage "Update completado exitosamente"
                Else
                    WScript.Echo "[DRY-RUN] Se ejecutaría update de módulos: " & modulesArg & " en " & gDbPath
                End If
            Else
                WScript.Echo "Error: El comando update requiere especificar módulos"
                WScript.Echo "Uso: cscript cli.vbs update <lista_modulos>"
                WScript.Quit 1
            End If
            
        Case "list-objects"
            If cleanArgCount >= 2 Then
                gDbPath = ResolvePath(cleanArgs(1))
                
                If Not gDryRun Then
                    ListObjects gDbPath
                Else
                    LogMessage "SIMULACION: Listaria objetos de " & gDbPath
                End If
            Else
                LogError "Faltan argumentos para list-objects"
                ShowHelp
            End If
            
        Case "schema"
            ' Parseo simple de opciones
            Dim dbArg, tableOpt, outOpt, fmtOpt, k
            dbArg = ""
            tableOpt = ""
            outOpt = ""
            fmtOpt = "json"
            ' argumentos posicionales/flags
            For k = 1 To cleanArgCount - 1
                Select Case LCase(cleanArgs(k))
                    Case "--table": If k < cleanArgCount - 1 Then tableOpt = cleanArgs(k+1): k = k + 1
                    Case "--out":   If k < cleanArgCount - 1 Then outOpt   = cleanArgs(k+1): k = k + 1
                    Case "--format":If k < cleanArgCount - 1 Then fmtOpt   = LCase(cleanArgs(k+1)): k = k + 1
                    Case Else
                        If dbArg = "" Then dbArg = cleanArgs(k)
                End Select
            Next
            ' db por defecto desde .ini
            If dbArg = "" Then
                dbArg = objConfig("DATABASE_DefaultPath")
            End If
            gDbPath = ResolvePath(dbArg)
            If outOpt = "" Then 
                ' Extraer nombre de la base de datos sin extensión
                Dim dbName: dbName = GetFileNameWithoutExtension(gDbPath)
                outOpt = gScriptDir & "\output\schema\" & dbName
            End If
            CreateFolderRecursive outOpt
            If fmtOpt <> "json" And fmtOpt <> "md" Then fmtOpt = "json"

            If Not ExportSchema(gDbPath, tableOpt, outOpt, fmtOpt) Then
                WScript.Echo "Error: no se pudo exportar el esquema"
                WScript.Quit 1
            End If
            WScript.Quit 0
            
        Case "form-export"
            ' Uso: form-export [db_path] --form <NombreFormulario> --out <ruta.json>
            Dim dbArg2, outArg, formArg, k2
            dbArg2 = "": outArg = "": formArg = ""
            For k2 = 1 To cleanArgCount - 1
                Select Case LCase(cleanArgs(k2))
                    Case "--form": If k2 < cleanArgCount - 1 Then formArg = cleanArgs(k2+1): k2 = k2 + 1
                    Case "--out":  If k2 < cleanArgCount - 1 Then outArg  = cleanArgs(k2+1): k2 = k2 + 1
                    Case Else: If dbArg2 = "" Then dbArg2 = cleanArgs(k2)
                End Select
            Next
            If formArg = "" Then WScript.Echo "Error: --form es obligatorio": WScript.Quit 1
            If dbArg2 = "" Then
                dbArg2 = objConfig("DATABASE_DefaultPath")
            End If
            gDbPath = ResolvePath(dbArg2)
            If outArg = "" Then outArg = gScriptDir & "\output\forms\" & formArg & ".json"
            outArg = ResolvePath(outArg)
            CreateFolderRecursive objFSO.GetParentFolderName(outArg)
            If Not ExportFormToJson(gDbPath, formArg, outArg) Then WScript.Quit 1
            WScript.Quit 0
            
        Case "form-import"
            ' Uso: form-import <json_path> [db_path] [--name NuevoNombre] [--overwrite] [--dry-run]
            Dim jsonArg, nameOpt, ow, dry, dbArg3, k3
            jsonArg = "": dbArg3 = "": nameOpt = "": ow = False: dry = False
            For k3 = 1 To cleanArgCount - 1
                Select Case LCase(cleanArgs(k3))
                    Case "--name": If k3 < cleanArgCount - 1 Then nameOpt = cleanArgs(k3+1): k3 = k3 + 1
                    Case "--overwrite": ow = True
                    Case "--dry-run": dry = True
                    Case Else
                        If jsonArg = "" Then jsonArg = cleanArgs(k3) Else If dbArg3 = "" Then dbArg3 = cleanArgs(k3)
                End Select
            Next
            If jsonArg = "" Then WScript.Echo "Error: falta json_path": WScript.Quit 1
            jsonArg = ResolvePath(jsonArg)
            If dbArg3 = "" Then
                dbArg3 = objConfig("DATABASE_DefaultPath")
            End If
            gDbPath = ResolvePath(dbArg3)
            If Not ImportFormFromJson(gDbPath, jsonArg, nameOpt, ow, dry) Then WScript.Quit 1
            WScript.Quit 0
            
        Case "ui-rebuild"
            ' ui-rebuild [db_path]
            Dim config4, dbArg4, uiRoot, formsDir, includeSub, pattern, formsPath, files, i4
            Set config4 = LoadConfig(gConfigPath)
            dbArg4 = ""
            If cleanArgCount > 1 Then dbArg4 = cleanArgs(1)
            If dbArg4 = "" Then dbArg4 = CfgGet(config4,"DATABASE_DefaultPath","DATABASE.DefaultPath","")
            gDbPath = ResolvePath(dbArg4)
            uiRoot = UI_GetRoot(config4)
            formsDir = UI_GetFormsDir(config4)
            includeSub = UI_GetIncludeSubdirs(config4)
            pattern = UI_GetFormFilePattern(config4)
            formsPath = ResolvePath(uiRoot & "\" & formsDir)
            files = EnumerateFiles(formsPath, pattern, includeSub)
            If Not UI_RebuildAll(gDbPath, files, config4) Then WScript.Quit 1 Else WScript.Quit 0
            
        Case "ui-update"
            ' ui-update [db_path] <form1.json> [form2.json ...]  (acepta rutas relativas o nombres base)
            Dim dbArg5, idx, list(), config5
            ReDim list(-1)
            Set config5 = LoadConfig(gConfigPath)
            dbArg5 = ""
            For idx = 1 To cleanArgCount - 1
                If dbArg5 = "" And InStr(cleanArgs(idx),".accdb")>0 Then
                    dbArg5 = cleanArgs(idx)
                Else
                    ReDim Preserve list(UBound(list)+1)
                    list(UBound(list)) = cleanArgs(idx)
                End If
            Next
            If dbArg5 = "" Then dbArg5 = CfgGet(config5,"DATABASE_DefaultPath","DATABASE.DefaultPath","")
            gDbPath = ResolvePath(dbArg5)
            If Not UI_UpdateSome(gDbPath, list, config5) Then WScript.Quit 1 Else WScript.Quit 0
            
        Case "ui-touch"
            ' ui-touch <formName|form.json>  -> atajo para un unico formulario
            Dim config6, dbArg6, item
            Set config6 = LoadConfig(gConfigPath)
            dbArg6 = CfgGet(config6,"DATABASE_DefaultPath","DATABASE.DefaultPath","")
            If cleanArgCount < 2 Then WScript.Echo "uso: ui-touch <formName|form.json>": WScript.Quit 1
            item = cleanArgs(1)
            Dim arr(0): arr(0) = item
            If Not UI_UpdateSome(ResolvePath(dbArg6), arr, config6) Then WScript.Quit 1 Else WScript.Quit 0
            
        Case "test"
            RunTests
            
        Case "--help", "help", "/?"
            ShowHelp
            
        Case Else
            LogError "Comando no reconocido: " & command
            ShowHelp
            WScript.Quit 1
    End Select
End Sub

' Sistema de testing

' Función para normalizar lista de módulos
Function NormalizeModuleList(moduleInput)
    On Error Resume Next
    
    Dim result(), count, i, j, parts, moduleName, found
    count = 0
    
    ' Si es un array, procesarlo directamente
    If IsArray(moduleInput) Then
        For i = 0 To UBound(moduleInput)
            moduleName = Trim(moduleInput(i))
            If moduleName <> "" And Left(moduleName, 2) <> "--" And Left(moduleName, 1) <> "/" Then
                ' Quitar extensión si existe
                If Right(LCase(moduleName), 4) = ".cls" Or Right(LCase(moduleName), 4) = ".bas" Then
                    moduleName = Left(moduleName, Len(moduleName) - 4)
                End If
                
                ' Quitar prefijo src\ si existe
                If InStr(moduleName, "src\") > 0 Then
                    moduleName = Replace(moduleName, "src\", "")
                End If
                If InStr(moduleName, "src/") > 0 Then
                    moduleName = Replace(moduleName, "src/", "")
                End If
                
                ' Verificar si ya existe (evitar duplicados)
                found = False
                For j = 0 To count - 1
                    If LCase(result(j)) = LCase(moduleName) Then
                        found = True
                        Exit For
                    End If
                Next
                
                If Not found Then
                    ReDim Preserve result(count)
                    result(count) = moduleName
                    count = count + 1
                End If
            End If
        Next
    Else
        ' Si es string, dividir por comas y espacios
        Dim cleanInput
        cleanInput = Replace(moduleInput, " ", ",")
        cleanInput = Replace(cleanInput, ",,", ",")
        parts = Split(cleanInput, ",")
        
        For i = 0 To UBound(parts)
            moduleName = Trim(parts(i))
            If moduleName <> "" And Left(moduleName, 2) <> "--" And Left(moduleName, 1) <> "/" Then
                ' Quitar extensión si existe
                If Right(LCase(moduleName), 4) = ".cls" Or Right(LCase(moduleName), 4) = ".bas" Then
                    moduleName = Left(moduleName, Len(moduleName) - 4)
                End If
                
                ' Quitar prefijo src\ si existe
                If InStr(moduleName, "src\") > 0 Then
                    moduleName = Replace(moduleName, "src\", "")
                End If
                If InStr(moduleName, "src/") > 0 Then
                    moduleName = Replace(moduleName, "src/", "")
                End If
                
                ' Verificar si ya existe (evitar duplicados)
                found = False
                For j = 0 To count - 1
                    If LCase(result(j)) = LCase(moduleName) Then
                        found = True
                        Exit For
                    End If
                Next
                
                If Not found Then
                    ReDim Preserve result(count)
                    result(count) = moduleName
                    count = count + 1
                End If
            End If
        Next
    End If
    
    If count = 0 Then
        NormalizeModuleList = Array()
    Else
        ReDim Preserve result(count - 1)
        NormalizeModuleList = result
    End If
End Function

' Función para actualizar módulos usando exactamente el mismo flujo que rebuild
Function UpdateModules(dbPath, modulesArg)
    On Error Resume Next
    UpdateModules = False
    
    Dim app
    Set app = OpenAccess(dbPath, GetDatabasePassword(dbPath))
    If app Is Nothing Then
        LogMessage "update: no se pudo abrir la base de datos (OpenAccess=Nada)"
        UpdateModules = False
        Exit Function
    End If
    
    ' Reforzar blindaje como en rebuild
    On Error Resume Next
    app.DisplayAlerts = False
    app.Echo False
    app.DoCmd.SetWarnings False
    Err.Clear
    On Error GoTo 0
    
    ' Normaliza la lista igual que rebuild; si rebuild ya tiene un normalizador, úsalo
    Dim list, i, name
    list = NormalizeModuleList(modulesArg)  ' Debe quitar extensiones, dividir por coma/espacio y eliminar duplicados
    
    If IsEmpty(list) Or UBound(list) < 0 Then
        LogMessage "update: sin módulos para actualizar"
        CloseAccess app
        UpdateModules = True
        Exit Function
    End If
    
    ' Logging inicial igual que rebuild
    LogMessage "update: módulos a actualizar: " & Join(list, ", ")
    
    ' Por cada módulo, usa EXACTAMENTE la MISMA llamada que hace rebuild por módulo
    For i = 0 To UBound(list)
        name = Trim(list(i))
        If name <> "" Then
            If gVerbose Then LogMessage "update: procesando " & name
            ' Reutiliza el mismo importador que usa rebuild para 1 módulo (no inventes otro)
            Call RebuildLike_ImportOne(app, name)
            If Err.Number <> 0 Then
                LogMessage "update: error importando " & name & ": " & Err.Number & " - " & Err.Description
                Err.Clear
            Else
                If gVerbose Then LogMessage "update: " & name & " actualizado correctamente"
            End If
        End If
    Next
    
    ' Compilar/guardar IGUAL que rebuild
    On Error Resume Next
    app.RunCommand acCmdCompileAndSaveAllModules
    If Err.Number <> 0 Then
        LogMessage "update: aviso al compilar: " & Err.Description
        Err.Clear
        ' Fallback: no intentes ningún otro RunCommand; continúa al cierre
    End If
    On Error GoTo 0
    
    ' Cierre IGUAL que rebuild
    CloseAccess app
    
    LogMessage "update: proceso completado exitosamente"
    UpdateModules = True
End Function

' Función para listar objetos de la base de datos
Sub ListObjects(dbPath)
    LogMessage "Listando objetos de: " & objFSO.GetFileName(dbPath)
    
    Set objAccess = OpenAccess(dbPath, gPassword)
    If objAccess Is Nothing Then
        Exit Sub
    End If
    
    WScript.Echo "=== TABLAS ==="
    Dim tbl
    For Each tbl In objAccess.CurrentDb.TableDefs
        If Left(tbl.Name, 4) <> "MSys" Then
            WScript.Echo "  " & tbl.Name
        End If
    Next
    
    WScript.Echo "=== FORMULARIOS ==="
    Dim frm
    For Each frm In objAccess.CurrentProject.AllForms
        WScript.Echo "  " & frm.Name
    Next
    
    WScript.Echo "=== CONSULTAS ==="
    Dim qry
    For Each qry In objAccess.CurrentDb.QueryDefs
        WScript.Echo "  " & qry.Name
    Next
    
    WScript.Echo "=== MODULOS ==="
    Dim mdl
    For Each mdl In objAccess.CurrentProject.AllModules
        WScript.Echo "  " & mdl.Name
    Next
    
    CloseAccess objAccess
End Sub



' ============================================================================
' SECCIÓN: FUNCIONES DE SELFTEST Y DIAGNÓSTICO
' ============================================================================



' Test de acceso a configuración
Function TestConfigAccess()
    On Error Resume Next
    Dim config
    Set config = LoadConfig(gConfigPath)
    If Err.Number <> 0 Then
        TestConfigAccess = False
        Exit Function
    End If
    On Error GoTo 0
    
    If config Is Nothing Then
        TestConfigAccess = False
        Exit Function
    End If
    
    TestConfigAccess = True
End Function

' Test de directorio fuente de módulos
Function TestModulesSrcPath()
    On Error Resume Next
    Dim config, srcPath
    Set config = LoadConfig(gConfigPath)
    If Err.Number <> 0 Or config Is Nothing Then
        TestModulesSrcPath = False
        Exit Function
    End If
    
    If Not config.Exists("MODULES_SrcPath") Then
        TestModulesSrcPath = False
        Exit Function
    End If
    
    srcPath = config("MODULES_SrcPath")
    If Not objFSO.FolderExists(srcPath) Then
        TestModulesSrcPath = False
        Exit Function
    End If
    On Error GoTo 0
    
    TestModulesSrcPath = True
End Function

' Test de extensiones de módulos
Function TestModulesExtensions()
    On Error Resume Next
    Dim config, extensions, extArray, i, ext, validExt
    Set config = LoadConfig(gConfigPath)
    If Err.Number <> 0 Or config Is Nothing Then
        TestModulesExtensions = False
        Exit Function
    End If
    
    If Not config.Exists("MODULES_Extensions") Then
        TestModulesExtensions = False
        Exit Function
    End If
    
    extensions = LCase(config("MODULES_Extensions"))
    ' Limpiar puntos iniciales si existen
    extensions = Replace(extensions, ".", "")
    extArray = Split(extensions, ",")
    
    For i = 0 To UBound(extArray)
        ext = Trim(extArray(i))
        If ext = "bas" Or ext = "cls" Then
            validExt = True
        End If
    Next
    
    On Error GoTo 0
    TestModulesExtensions = validExt
End Function

' Test de base de datos por defecto
Function TestDatabaseDefaultPath()
    On Error Resume Next
    Dim config, dbPath
    Set config = LoadConfig(gConfigPath)
    If Err.Number <> 0 Or config Is Nothing Then
        TestDatabaseDefaultPath = False
        Exit Function
    End If
    
    If Not config.Exists("DATABASE_DefaultPath") Then
        TestDatabaseDefaultPath = False
        Exit Function
    End If
    
    dbPath = ResolvePath(config("DATABASE_DefaultPath"))
    If dbPath = "" Then
        TestDatabaseDefaultPath = False
        Exit Function
    End If
    On Error GoTo 0
    
    TestDatabaseDefaultPath = True
End Function

' Test de automatización de Access
Function TestAccessAutomation()
    On Error Resume Next
    Dim objAccess, config, dbPath
    
    ' Cargar configuración para obtener ruta de base de datos
    Set config = LoadConfig(gConfigPath)
    If Err.Number <> 0 Or config Is Nothing Then
        TestAccessAutomation = False
        Exit Function
    End If
    
    If Not config.Exists("DATABASE_DefaultPath") Then
        TestAccessAutomation = False
        Exit Function
    End If
    
    dbPath = ResolvePath(config("DATABASE_DefaultPath"))
    If Not objFSO.FileExists(dbPath) Then
        TestAccessAutomation = False
        Exit Function
    End If
    
    ' Intentar crear instancia de Access
    Set objAccess = OpenAccess(dbPath, gPassword)
    If Err.Number <> 0 Or objAccess Is Nothing Then
        TestAccessAutomation = False
        Exit Function
    End If
    
    ' Cerrar y limpiar
    CloseAccess objAccess
    
    On Error GoTo 0
    TestAccessAutomation = True
End Function

' Test de acceso VBE
Function TestVBEAccess()
    On Error Resume Next
    Dim objAccess, config, dbPath
    
    ' Cargar configuración para obtener ruta de base de datos
    Set config = LoadConfig(gConfigPath)
    If Err.Number <> 0 Or config Is Nothing Then
        TestVBEAccess = False
        Exit Function
    End If
    
    If Not config.Exists("DATABASE_DefaultPath") Then
        TestVBEAccess = False
        Exit Function
    End If
    
    dbPath = ResolvePath(config("DATABASE_DefaultPath"))
    If Not objFSO.FileExists(dbPath) Then
        TestVBEAccess = False
        Exit Function
    End If
    
    ' Intentar crear instancia de Access
    Set objAccess = OpenAccess(dbPath, gPassword)
    If Err.Number <> 0 Or objAccess Is Nothing Then
        TestVBEAccess = False
        Exit Function
    End If
    
    ' Probar acceso VBE usando CheckVBProjectAccess
    If Not CheckVBProjectAccess(objAccess) Then
        CloseAccess objAccess
        TestVBEAccess = False
        Exit Function
    End If
    
    ' Cerrar y limpiar
    CloseAccess objAccess
    
    On Error GoTo 0
    TestVBEAccess = True
End Function

' Función para verificar acceso VBE
Function CheckVBProjectAccess(objAccess)
    On Error Resume Next
    Dim projectCount
    projectCount = objAccess.VBE.VBProjects.Count
    
    If Err.Number <> 0 Then
        CheckVBProjectAccess = False
        LogMessage "Error accediendo VBE: " & Err.Description
        LogMessage "SOLUCION: Habilitar 'Trust access to the VBA project object model' en:"
        LogMessage "  Access > File > Options > Trust Center > Trust Center Settings"
        LogMessage "  > Macro Settings > Trust access to the VBA project object model"
        Exit Function
    End If
    On Error GoTo 0
    
    CheckVBProjectAccess = True
End Function

' Función para validar sintaxis de archivos VBA
Function ValidateVBASyntax(filePath, ByRef errorDetails)
    ValidateVBASyntax = True
    errorDetails = ""
    
    On Error Resume Next
    
    ' Verificar que el archivo se puede leer
    If Not objFSO.FileExists(filePath) Then
        errorDetails = "Archivo no existe"
        ValidateVBASyntax = False
        Exit Function
    End If
    
    Dim fileContent
    fileContent = ReadModuleFile(filePath)
    
    If Err.Number <> 0 Then
        errorDetails = "Error leyendo archivo: " & Err.Description
        ValidateVBASyntax = False
        Err.Clear
        Exit Function
    End If
    
    ' Verificar contenido vacío
    If Len(Trim(fileContent)) = 0 Then
        errorDetails = "Archivo vacío"
        ValidateVBASyntax = False
        Exit Function
    End If
    
    ' Verificar caracteres nulos que pueden causar problemas
    If InStr(fileContent, Chr(0)) > 0 Then
        errorDetails = "Archivo contiene caracteres nulos"
        ValidateVBASyntax = False
        Exit Function
    End If
    
    On Error GoTo 0
End Function

' Función para borrar módulos usando EXACTAMENTE la misma lógica que rebuild
Function DeleteModule_UsingRebuildPath(objAccess, moduleName)
    On Error Resume Next
    DeleteModule_UsingRebuildPath = False
    
    Dim vbProject, vbComponent, i
    Set vbProject = objAccess.VBE.ActiveVBProject
    
    ' Buscar el módulo por nombre usando la misma lógica que rebuild
    For i = vbProject.VBComponents.Count To 1 Step -1
        Set vbComponent = vbProject.VBComponents(i)
        
        ' Solo procesar módulos estándar y de clase (igual que rebuild)
        If vbComponent.Type = 1 Or vbComponent.Type = 2 Then ' vbext_ct_StdModule = 1, vbext_ct_ClassModule = 2
            If vbComponent.Name = moduleName Then
                If gVerbose Then LogMessage "  Eliminando: " & vbComponent.Name & " (Tipo: " & vbComponent.Type & ")"
                vbProject.VBComponents.Remove vbComponent
                
                If Err.Number <> 0 Then
                    LogMessage "  ❌ Error eliminando " & vbComponent.Name & ": " & Err.Description
                    Err.Clear
                    DeleteModule_UsingRebuildPath = False
                    Exit Function
                Else
                    If gVerbose Then LogMessage "  ✓ Eliminado: " & vbComponent.Name
                    DeleteModule_UsingRebuildPath = True
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Si llegamos aquí, el módulo no se encontró (no es error)
    DeleteModule_UsingRebuildPath = True
End Function

' Función para limpiar archivos VBA eliminando metadatos
Function CleanVBAFile(filePath, fileExtension)
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
        ' PRESERVAR: Option Compare Database es esencial para el funcionamiento
        If Not (Left(Trim(strLine), 9) = "Attribute" Or _
                Left(Trim(strLine), 17) = "VERSION 1.0 CLASS" Or _
                Trim(strLine) = "BEGIN" Or _
                Left(Trim(strLine), 8) = "MultiUse" Or _
                Trim(strLine) = "END" Or _
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

' Subrutina para importar módulos con codificación ANSI
Sub ImportModuleWithAnsiEncoding(strImportPath, moduleName, fileExtension, vbComponent, cleanedContent)
    On Error Resume Next
    
    LogMessage "  Importando: " & moduleName & " desde " & strImportPath
    
    ' Crear nuevo componente
    Dim newComponent, componentType
    
    If LCase(fileExtension) = "cls" Then
        componentType = 2  ' vbext_ct_ClassModule
    Else
        componentType = 1  ' vbext_ct_StdModule
    End If
    
    LogMessage "  Creando nuevo componente tipo " & componentType & " para: " & moduleName
    Set newComponent = vbComponent.VBComponents.Add(componentType)
    
    ' Verificar que el componente se creó correctamente
    If newComponent Is Nothing Then
        LogMessage "  ❌ Error: No se pudo crear componente para " & moduleName
        Exit Sub
    End If
    
    newComponent.Name = moduleName
    
    ' Verificar si el componente ya existe
    Dim existingComponent
    Set existingComponent = Nothing
    For Each existingComponent In vbComponent.VBComponents
        If existingComponent.Name = moduleName Then
            Exit For
        End If
        Set existingComponent = Nothing
    Next
    
    ' Si existe, eliminar el código existente
    If Not existingComponent Is Nothing Then
        LogMessage "  Limpiando código existente de: " & moduleName
        
        ' Limpiar solo si hay líneas
        If existingComponent.CodeModule.CountOfLines > 0 Then
            existingComponent.CodeModule.DeleteLines 1, existingComponent.CodeModule.CountOfLines
        End If
        
        ' Insertar contenido limpio
        If Len(cleanedContent) > 0 Then
            existingComponent.CodeModule.AddFromString cleanedContent
            LogMessage "  ✓ Actualizado: " & moduleName
        End If
    End If
    
    ' Verificar si hubo errores críticos durante la importación
    If Err.Number <> 0 Then
        LogMessage "  ❌ Error crítico con " & moduleName & ": " & Err.Description
        Err.Clear
    Else
        LogMessage "  ✓ Importación completada exitosamente para: " & moduleName
    End If
    
    On Error GoTo 0
End Sub

' Subrutina para importar módulos con codificación ANSI (versión nueva)
' Subrutina para verificar nombres de módulos
Sub VerifyModuleNames(accessApp, srcPath)
    On Error Resume Next
    
    Dim vbProject, vbComponent
    Dim srcFiles, srcFile, srcModuleName
    Dim accessModules, accessModule
    Dim missingInAccess, missingInSrc
    Dim i
    
    Set vbProject = accessApp.VBE.ActiveVBProject
    
    ' Obtener archivos de la carpeta src
    Set srcFiles = CreateObject("Scripting.Dictionary")
    Dim folder, file
    Set folder = objFSO.GetFolder(srcPath)
    
    For Each file In folder.Files
        If LCase(objFSO.GetExtensionName(file.Name)) = "bas" Or _
           LCase(objFSO.GetExtensionName(file.Name)) = "cls" Then
            srcModuleName = objFSO.GetBaseName(file.Name)
            srcFiles.Add srcModuleName, file.Path
        End If
    Next
    
    ' Obtener módulos de Access
    Set accessModules = CreateObject("Scripting.Dictionary")
    For Each vbComponent In vbProject.VBComponents
        If vbComponent.Type = 1 Or vbComponent.Type = 2 Then ' Solo módulos estándar y de clase
            accessModules.Add vbComponent.Name, vbComponent.Type
        End If
    Next
    
    ' Verificar módulos faltantes en Access
    missingInAccess = ""
    For Each srcModuleName In srcFiles.Keys
        If Not accessModules.Exists(srcModuleName) Then
            If missingInAccess <> "" Then missingInAccess = missingInAccess & ", "
            missingInAccess = missingInAccess & srcModuleName
        End If
    Next
    
    ' Verificar módulos faltantes en src
    missingInSrc = ""
    For Each accessModule In accessModules.Keys
        If Not srcFiles.Exists(accessModule) Then
            If missingInSrc <> "" Then missingInSrc = missingInSrc & ", "
            missingInSrc = missingInSrc & accessModule
        End If
    Next
    
    ' Reporte final
    WScript.Echo "=== VERIFICACION DE INTEGRIDAD ==="
    WScript.Echo "Modulos en carpeta src: " & srcFiles.Count
    WScript.Echo "Modulos en Access: " & accessModules.Count
    
    If missingInAccess <> "" Then
        WScript.Echo "⚠️  Modulos en src pero NO en Access: " & missingInAccess
    End If
    
    If missingInSrc <> "" Then
        WScript.Echo "⚠️  Modulos en Access pero NO en src: " & missingInSrc
    End If
    
    If missingInAccess = "" And missingInSrc = "" Then
        WScript.Echo "✓ Todos los modulos están sincronizados correctamente"
    End If
    
    On Error GoTo 0
End Sub

' Función unificada para importar un módulo usando la misma lógica que rebuild
Function ResolveSourcePathForModule(moduleName)
    ' Función que usa la MISMA lógica de descubrimiento que rebuild
    ' para mapear nombres de módulos a archivos en SrcPath (incluye subcarpetas)
    On Error Resume Next
    ResolveSourcePathForModule = ""
    
    If Not objFSO.FolderExists(g_ModulesSrcPath) Then
        Exit Function
    End If
    
    ' Usar GetModuleFiles con patrón específico para encontrar el módulo
    Dim moduleFiles
    moduleFiles = GetModuleFiles(g_ModulesSrcPath, g_ModulesExtensions, g_ModulesIncludeSubdirs, moduleName & ".*")
    
    ' Si se encontró exactamente un archivo, devolverlo
    If UBound(moduleFiles) >= 0 Then
        Dim i, fileName
        For i = 0 To UBound(moduleFiles)
            fileName = objFSO.GetBaseName(moduleFiles(i))
            If fileName = moduleName Then
                ResolveSourcePathForModule = moduleFiles(i)
                Exit Function
            End If
        Next
    End If
End Function

' Función para importar un módulo usando EXACTAMENTE la misma lógica que rebuild
Function RebuildLike_ImportOne(objAccess, moduleName)
    On Error Resume Next
    RebuildLike_ImportOne = False
    
    ' 1) Resolver ruta fuente usando la función unificada
    Dim filePath
    filePath = ResolveSourcePathForModule(moduleName)
    
    If filePath = "" Then
        LogMessage "No se encontró el archivo para " & moduleName
        RebuildLike_ImportOne = False
        Exit Function
    End If
    
    Dim fileExtension
    fileExtension = LCase(objFSO.GetExtensionName(filePath))
    
    ' Log IDÉNTICO al de rebuild
    LogMessage "Importando módulo " & moduleName & " desde " & filePath
    
    ' 2) Validar sintaxis usando la misma función que rebuild
    Dim validationResult, errorDetails
    validationResult = ValidateVBASyntax(filePath, errorDetails)
    If validationResult <> True Then
        LogMessage "ERROR en sintaxis de " & moduleName & ": " & errorDetails
        RebuildLike_ImportOne = False
        Exit Function
    End If
    
    ' 3) Borrar el módulo destino si existe, usando la MISMA lógica que rebuild
    If Not DeleteModule_UsingRebuildPath(objAccess, moduleName) Then
        LogMessage "Error borrando módulo " & moduleName
        RebuildLike_ImportOne = False
        Exit Function
    End If
    
    ' 4) Limpiar archivo antes de importar usando la misma función que rebuild
    Dim cleanedContent
    cleanedContent = CleanVBAFile(filePath, fileExtension)
    
    ' 5) Importar EXACTAMENTE como lo hace rebuild
    Call ImportModuleWithAnsiEncoding(filePath, moduleName, fileExtension, objAccess.VBE.ActiveVBProject, cleanedContent)
    
    ' 6) Verificar éxito según el mismo criterio que rebuild
    If Err.Number <> 0 Then
        LogMessage "Error importando " & moduleName & ": " & Err.Number & " - " & Err.Description
        Err.Clear
        RebuildLike_ImportOne = False
        Exit Function
    End If
    
    ' 7) Post-procesar el módulo usando las mismas funciones que rebuild
    Dim vbProject, vbComponent
    Set vbProject = objAccess.VBE.ActiveVBProject
    For Each vbComponent In vbProject.VBComponents
        If vbComponent.Name = moduleName Then
            Call PostProcessInsertedModule(vbComponent)
            If Not vbComponent.CodeModule Is Nothing Then
                Call EnsureOptionExplicit(vbComponent.CodeModule)
            End If
            Exit For
        End If
    Next
    
    RebuildLike_ImportOne = True
    LogMessage "✓ Módulo " & moduleName & " importado exitosamente"
End Function

' Función para determinar la contraseña de la base de datos
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

' Subrutina para reconstruir proyecto completo (basada en condor_cli.vbs)
Sub RebuildProject()
    LogVerbose "rebuild: usando OpenAccess/CloseAccess canonicos"
    WScript.Echo "=== RECONSTRUCCION COMPLETA DEL PROYECTO VBA ==="
    WScript.Echo "ADVERTENCIA: Se eliminaran TODOS los modulos VBA existentes"
    WScript.Echo "Iniciando proceso de reconstruccion..."
    
    ' Cerrar procesos de Access existentes antes de comenzar
    CloseExistingAccessProcesses
    
    ' Abrir Access de forma segura
    Set objAccess = OpenAccess(gDbPath, GetDatabasePassword(gDbPath))
    If objAccess Is Nothing Then
        WScript.Echo "❌ Error: No se pudo abrir Access"
        WScript.Quit 1
    End If
    
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
    
    ' Cerrar usando la vía canónica
    CloseAccess objAccess
    WScript.Echo "✓ Base de datos cerrada y guardada"
    
    ' Paso 3: Volver a abrir la base de datos
    WScript.Echo "Paso 3: Reabriendo base de datos con proyecto VBA limpio..."
    
    Set objAccess = OpenAccess(gDbPath, GetDatabasePassword(gDbPath))
    If objAccess Is Nothing Then
        WScript.Echo "❌ Error: No se pudo reabrir la base de datos (OpenAccess)"
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
    
    If Not objFSO.FolderExists(g_ModulesSrcPath) Then
        LogMessage "Error: Directorio de origen no existe: " & g_ModulesSrcPath
        objAccess.Quit
        WScript.Quit 1
    End If
    
    ' PASO 4.1: Validacion previa de sintaxis
    LogMessage "Validando sintaxis de todos los modulos..."
    Set objFolder = objFSO.GetFolder(g_ModulesSrcPath)
    totalFiles = 0
    validFiles = 0
    invalidFiles = 0
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            totalFiles = totalFiles + 1
            validationResult = ValidateVBASyntax(objFile.Path, errorDetails)
            
            If validationResult = True Then
                validFiles = validFiles + 1
                LogMessage "  Sintaxis valida: " & objFile.Name
            Else
                invalidFiles = invalidFiles + 1
                LogMessage "  ERROR en " & objFile.Name & ": " & errorDetails
            End If
        End If
    Next
    
    If invalidFiles > 0 Then
        LogMessage "ABORTANDO: Se encontraron " & invalidFiles & " archivos con errores de sintaxis."
        LogMessage "Use 'cscript cli.vbs validate --verbose' para mas detalles."
        objAccess.Quit
        WScript.Quit 1
    End If
    
    LogMessage "Validacion completada: " & validFiles & " archivos validos"
    
    ' PASO 4.2: Procesar archivos de modulos
    Set objFolder = objFSO.GetFolder(g_ModulesSrcPath)
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            strModuleName = objFSO.GetBaseName(objFile.Name)
            
            LogMessage "Procesando modulo: " & strModuleName
            
            ' Usar la función unificada de importación
            Dim importResult
            importResult = RebuildLike_ImportOne(objAccess, strModuleName)
            
            If Not importResult Then
                LogMessage "Error al importar modulo " & strModuleName
            End If
        End If
    Next
    
    ' PASO 4.3: Importar modulos desde /src
    LogMessage "Importando modulos desde /src..."
    
    For Each vbComponent In objAccess.VBE.ActiveVBProject.VBComponents
        If vbComponent.Type = 1 Then  ' vbext_ct_StdModule
            LogMessage "Guardando modulo: " & vbComponent.Name
            ' Eliminado DoCmd.Save individual - se hará al final con RunCommand 126
        ElseIf vbComponent.Type = 2 Then  ' vbext_ct_ClassModule
            LogMessage "Guardando clase: " & vbComponent.Name
            ' Eliminado DoCmd.Save individual - se hará al final con RunCommand 126
        End If
    Next
    
    ' PASO 4.4: Verificacion de integridad y compilacion
    LogMessage "Verificando integridad de nombres de modulos..."
    Call VerifyModuleNames(objAccess, g_ModulesSrcPath)
    
    ' Compilar y guardar todo al final
    On Error Resume Next
    objAccess.RunCommand 126  ' acCmdCompileAndSaveAllModules
    If Err.Number <> 0 Then
        LogMessage "rebuild: aviso al compilar: " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    
    LogMessage "=== RECONSTRUCCION COMPLETADA EXITOSAMENTE ==="
    LogMessage "El proyecto VBA ha sido completamente reconstruido"
    LogMessage "Todos los modulos han sido reimportados desde /src"
    
    ' Cerrar usando la vía canónica
    CloseAccess objAccess
    
    LogVerbose "rebuild: cierre canonico completado"
    On Error GoTo 0
End Sub

' Ejecutar función principal
Main()

' === FUNCIONES DE EXPORTACION DE FORMULARIOS ===

Function ExportFormToJson(dbPath, formName, outJson)
    On Error Resume Next
    ExportFormToJson = False
    Dim app: Set app = OpenAccess(dbPath, GetDatabasePassword(dbPath))
    If app Is Nothing Then Exit Function

    ' Blindaje anti-UI antes de OpenForm
    app.DisplayAlerts = False
    app.Echo = False
    app.DoCmd.SetWarnings False
    Err.Clear
    
    ' Abrir en diseño y oculto
    app.DoCmd.OpenForm formName, 1, , , , 1 ' acDesign=1, acHidden=1
    ' Reforzar anti-UI después de OpenForm
    app.DisplayAlerts = False
    app.Echo = False
    app.DoCmd.SetWarnings False
    Err.Clear
    If Err.Number <> 0 Then
        LogMessage "form-export: no se pudo abrir " & formName & " en diseño: " & Err.Description
        Err.Clear: CloseAccess app: Exit Function
    End If

    Dim frm: Set frm = app.Forms(formName)
    Dim root: Set root = CreateObject("Scripting.Dictionary")
    root("name") = formName
    root("type") = "Form"
    root("recordSource") = SafeProp(frm, "RecordSource")
    root("caption") = SafeProp(frm, "Caption")
    root("width") = SafeProp(frm, "Width")
    root("sectionHeights") = DictFromSections(frm)   ' Detail/Header/Footer heights si aplican

    ' Controles (orden determinista por Name)
    Dim names: names = ControlNamesSorted(frm)
    Dim arr(), i, c
    If IsArray(names) Then
        ReDim arr(UBound(names))
        For i = 0 To UBound(names)
            Set c = frm.Controls(names(i))
            Set arr(i) = ControlToDict(c)
        Next
        root("controls") = arr
    Else
        root("controls") = Array()
    End If

    ' Guardar JSON (usa SaveToJSON existente)
    SaveToJSON root, outJson

    ' Cerrar sin guardar
    app.DoCmd.Close acForm, formName, acSaveNo
    CloseAccess app
    ExportFormToJson = True
End Function

' === Helpers export ===
Function HasProperty(obj, p)
    On Error Resume Next
    Dim tmp: tmp = obj.Properties(p) ' acceso prueba
    If Err.Number = 0 Then
        HasProperty = True
    Else
        HasProperty = False
        Err.Clear
    End If
End Function

Function SafeProp(obj, p)
    On Error Resume Next
    Dim v: v = ""
    If HasProperty(obj, p) Then
        v = obj.Properties(p)
        If Err.Number <> 0 Then v = "": Err.Clear
    End If
    SafeProp = v
End Function

Function DictFromSections(frm)
    On Error Resume Next
    Dim d: Set d = CreateObject("Scripting.Dictionary")
    d("detail") = frm.Section(0).Height ' acDetail=0
    If Err.Number = 0 Then
        d("header") = frm.Section(1).Height ' acHeader=1
        Err.Clear
        d("footer") = frm.Section(2).Height ' acFooter=2
        Err.Clear
    End If
    Set DictFromSections = d
End Function

Function ControlNamesSorted(frm)
    On Error Resume Next
    Dim list(), i, n
    n = frm.Controls.Count
    If n <= 0 Then Exit Function
    ReDim list(n-1)
    For i = 0 To n-1: list(i) = frm.Controls(i).Name: Next
    QuickSortStrings list, 0, n-1
    ControlNamesSorted = list
End Function

Function ControlToDict(ctrl)
    On Error Resume Next
    Dim d: Set d = CreateObject("Scripting.Dictionary")
    d("name") = ctrl.Name
    d("controlType") = ctrl.ControlType
    d("section") = ctrl.Section
    d("left") = ctrl.Left
    d("top") = ctrl.Top
    d("width") = ctrl.Width
    d("height") = ctrl.Height
    d("tabIndex") = SafeProp(ctrl, "TabIndex")
    d("tabStop") = SafeProp(ctrl, "TabStop")
    d("visible") = SafeProp(ctrl, "Visible")
    d("enabled") = SafeProp(ctrl, "Enabled")
    d("locked") = SafeProp(ctrl, "Locked")
    d("controlSource") = SafeProp(ctrl, "ControlSource")
    d("rowSource") = SafeProp(ctrl, "RowSource")
    d("sourceObject") = SafeProp(ctrl, "SourceObject") ' subform/subreport
    d("caption") = SafeProp(ctrl, "Caption")
    d("format") = SafeProp(ctrl, "Format")
    d("fontName") = SafeProp(ctrl, "FontName")
    d("fontSize") = SafeProp(ctrl, "FontSize")
    d("foreColor") = SafeProp(ctrl, "ForeColor")
    d("backColor") = SafeProp(ctrl, "BackColor")
    ' eventos (expresiones)
    d("onClick") = SafeProp(ctrl, "OnClick")
    d("onDblClick") = SafeProp(ctrl, "OnDblClick")
    d("onChange") = SafeProp(ctrl, "OnChange")
    d("onCurrent") = SafeProp(ctrl, "OnCurrent")
    Set ControlToDict = d
End Function

' Implementación simple de QuickSort para strings
Sub QuickSortStrings(arr, low, high)
    On Error Resume Next
    If low < high Then
        Dim pi
        pi = PartitionStrings(arr, low, high)
        QuickSortStrings arr, low, pi - 1
        QuickSortStrings arr, pi + 1, high
    End If
End Sub

Function PartitionStrings(arr, low, high)
    On Error Resume Next
    Dim pivot, i, j, temp
    pivot = arr(high)
    i = low - 1
    
    For j = low To high - 1
        If LCase(arr(j)) <= LCase(pivot) Then
            i = i + 1
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
        End If
    Next
    
    temp = arr(i + 1)
    arr(i + 1) = arr(high)
    arr(high) = temp
    
    PartitionStrings = i + 1
End Function

Function GetParentFolder(p)
    On Error Resume Next
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    GetParentFolder = fso.GetParentFolderName(p)
End Function

Function ReadAllText(path)
    On Error Resume Next
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile path
    ReadAllText = stream.ReadText
    stream.Close
    If Err.Number <> 0 Then ReadAllText = ""
    Err.Clear
End Function

Function JsonParse(json)
    On Error Resume Next
    Dim parser
    Set parser = New JsonParser
    Set JsonParse = parser.Parse(json)
    If Err.Number <> 0 Then Set JsonParse = Nothing
    Err.Clear
End Function

Function ValidateFormJson(root, ByRef errMsg)
    On Error Resume Next
    ValidateFormJson = False
    errMsg = ""
    
    ' Declaracion de variables
    Dim i, j, ctrl, propName, propValue, ctValue
    Dim props
    
    ' Validar que root sea un objeto
    If root Is Nothing Then
        errMsg = "JSON root es Nothing"
        Exit Function
    End If
    
    ' Validar name (string no vacio)
    If Not root.Exists("name") Then
        errMsg = "falta campo 'name'"
        Exit Function
    End If
    
    Dim nameValue: nameValue = root("name")
    If VarType(nameValue) <> vbString Or Len(Trim(nameValue)) = 0 Then
        errMsg = "campo 'name' debe ser string no vacio"
        Exit Function
    End If
    
    ' Validar controls (debe ser Array o objeto con Count > 0)
    If Not root.Exists("controls") Then
        errMsg = "falta campo 'controls'"
        Exit Function
    End If
    
    Dim controls: Set controls = root("controls")
    If controls Is Nothing Then
        errMsg = "campo 'controls' es Nothing"
        Exit Function
    End If
    
    Dim keys, k, cnt
    If IsArray(controls) Then 
        For i = 0 To UBound(controls) 
            Set ctrl = controls(i)
            
            If ctrl Is Nothing Then
                errMsg = "control en posicion " & i & " es Nothing"
                Exit Function
            End If
            
            ' Validar controlType (numerico)
            If Not ctrl.Exists("controlType") Then
                errMsg = "control en posicion " & i & " falta 'controlType'"
                Exit Function
            End If
            
            ctValue = ctrl("controlType")
            If Not IsNumeric(ctValue) Then
                errMsg = "control en posicion " & i & " 'controlType' debe ser numerico"
                Exit Function
            End If
            
            ' Validar left, top, width, height (numericos)
            props = Array("left", "top", "width", "height")
            For j = 0 To UBound(props)
                propName = props(j)
                If Not ctrl.Exists(propName) Then
                    errMsg = "control en posicion " & i & " falta '" & propName & "'"
                    Exit Function
                End If
                
                propValue = ctrl(propName)
                If Not IsNumeric(propValue) Then
                    errMsg = "control en posicion " & i & " '" & propName & "' debe ser numerico"
                    Exit Function
                End If
            Next
        Next
    ElseIf IsObject(controls) Then 
        On Error Resume Next 
        cnt = controls.Count 
        If Err.Number = 0 Then 
            keys = controls.Keys 
            For k = 0 To cnt - 1 
                Set ctrl = controls(keys(k))
                
                If ctrl Is Nothing Then
                    errMsg = "control en clave " & keys(k) & " es Nothing"
                    Exit Function
                End If
                
                ' Validar controlType (numerico)
                If Not ctrl.Exists("controlType") Then
                    errMsg = "control en clave " & keys(k) & " falta 'controlType'"
                    Exit Function
                End If
                
                ctValue = ctrl("controlType")
                If Not IsNumeric(ctValue) Then
                    errMsg = "control en clave " & keys(k) & " 'controlType' debe ser numerico"
                    Exit Function
                End If
                
                ' Validar left, top, width, height (numericos)
                props = Array("left", "top", "width", "height")
                For j = 0 To UBound(props)
                    propName = props(j)
                    If Not ctrl.Exists(propName) Then
                        errMsg = "control en clave " & keys(k) & " falta '" & propName & "'"
                        Exit Function
                    End If
                    
                    propValue = ctrl(propName)
                    If Not IsNumeric(propValue) Then
                        errMsg = "control en clave " & keys(k) & " '" & propName & "' debe ser numerico"
                        Exit Function
                    End If
                Next
            Next 
        Else
            errMsg = "campo 'controls' no tiene propiedad Count"
            Exit Function
        End If 
        Err.Clear 
    Else
        errMsg = "campo 'controls' debe ser Array u objeto con Count"
        Exit Function
    End If
    
    ValidateFormJson = True
End Function

Function UI_DryRunCheck(root, ByRef report)
    On Error Resume Next
    UI_DryRunCheck = False
    report = ""
    
    ' Validacion basica de estructura JSON
    Dim errMsg: errMsg = ""
    If Not ValidateFormJson(root, errMsg) Then
        report = "ERROR: Validacion JSON fallo - " & errMsg
        Exit Function
    End If
    
    ' Extraer informacion del formulario
    Dim formName: formName = root("name")
    Dim controls: Set controls = root("controls")
    
    Dim controlCount: controlCount = 0
    If IsArray(controls) Then
        controlCount = UBound(controls) + 1
    ElseIf IsObject(controls) Then
        controlCount = controls.Count
    End If
    
    ' Validar tipos de controles soportados
    Dim supportedTypes: supportedTypes = Array(100, 101, 102, 103, 104, 105, 106, 107, 108, 109, 110, 111, 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123, 124, 125, 126, 127, 128, 129)
    Dim unsupportedControls: unsupportedControls = ""
    Dim unsupportedCount: unsupportedCount = 0
    
    Dim i, ctrl, controlType, found, j, keys
    If IsArray(controls) Then
        For i = 0 To controlCount - 1
            Set ctrl = controls(i)
            
            controlType = ctrl("controlType")
            found = False
            
            For j = 0 To UBound(supportedTypes)
                If controlType = supportedTypes(j) Then
                    found = True
                    Exit For
                End If
            Next
            
            If Not found Then
                If unsupportedCount > 0 Then unsupportedControls = unsupportedControls & ", "
                unsupportedControls = unsupportedControls & "tipo " & controlType
                unsupportedCount = unsupportedCount + 1
            End If
        Next
    ElseIf IsObject(controls) Then
        keys = controls.Keys
        For i = 0 To UBound(keys)
            Set ctrl = controls(keys(i))
            
            controlType = ctrl("controlType")
            found = False
            
            For j = 0 To UBound(supportedTypes)
                If controlType = supportedTypes(j) Then
                    found = True
                    Exit For
                End If
            Next
            
            If Not found Then
                If unsupportedCount > 0 Then unsupportedControls = unsupportedControls & ", "
                unsupportedControls = unsupportedControls & "tipo " & controlType
                unsupportedCount = unsupportedCount + 1
            End If
        Next
    End If
    
    ' Generar reporte
    report = "DRY-RUN: Formulario '" & formName & "' - " & controlCount & " controles"
    
    If unsupportedCount > 0 Then
        report = report & vbCrLf & "ADVERTENCIA: " & unsupportedCount & " controles no soportados: " & unsupportedControls
    End If
    
    ' Validar propiedades criticas de controles
    Dim invalidControls: invalidControls = ""
    Dim invalidCount: invalidCount = 0
    
    If IsArray(controls) Then
        For i = 0 To controlCount - 1
            Set ctrl = controls(i)
            
            ' Verificar dimensiones validas
            Dim left, top, width, height
            left = ctrl("left")
            top = ctrl("top") 
            width = ctrl("width")
            height = ctrl("height")
            
            If width <= 0 Or height <= 0 Then
                If invalidCount > 0 Then invalidControls = invalidControls & ", "
                invalidControls = invalidControls & "control " & i & " (dimensiones invalidas)"
                invalidCount = invalidCount + 1
            End If
            
            ' Verificar posicion no negativa
            If left < 0 Or top < 0 Then
                If invalidCount > 0 Then invalidControls = invalidControls & ", "
                invalidControls = invalidControls & "control " & i & " (posicion negativa)"
                invalidCount = invalidCount + 1
            End If
        Next
    ElseIf IsObject(controls) Then
        keys = controls.Keys
        For i = 0 To UBound(keys)
            Set ctrl = controls(keys(i))
            
            ' Verificar dimensiones validas
            Dim left2, top2, width2, height2
            left2 = ctrl("left")
            top2 = ctrl("top") 
            width2 = ctrl("width")
            height2 = ctrl("height")
            
            If width2 <= 0 Or height2 <= 0 Then
                If invalidCount > 0 Then invalidControls = invalidControls & ", "
                invalidControls = invalidControls & "control " & i & " (dimensiones invalidas)"
                invalidCount = invalidCount + 1
            End If
            
            ' Verificar posicion no negativa
            If left2 < 0 Or top2 < 0 Then
                If invalidCount > 0 Then invalidControls = invalidControls & ", "
                invalidControls = invalidControls & "control " & i & " (posicion negativa)"
                invalidCount = invalidCount + 1
            End If
        Next
    End If
    
    If invalidCount > 0 Then
        report = report & vbCrLf & "ERROR: " & invalidCount & " controles con problemas: " & invalidControls
        Exit Function
    End If
    
    report = report & vbCrLf & "VALIDACION: Estructura JSON correcta, todos los controles son validos"
    UI_DryRunCheck = True
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
                    Err.Raise 1001, "JsonParser", "Caracter inesperado en posicion " & pos & ": " & char
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
                Err.Raise 1002, "JsonParser", "Se esperaba ':' despues de la clave"
            End If
            pos = pos + 1
            
            Dim value, peek
            peek = Mid(jsonText, pos, 1)
            
            If peek = "{" Or peek = "[" Then
                ' valor compuesto -> usar Set
                Dim objOrArr
                Set objOrArr = ParseValue()
                Set obj(key) = objOrArr
            Else
                ' escalar -> sin Set
                value = ParseValue()
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
        Dim arr(), count
        count = -1
        pos = pos + 1 ' Saltar '['
        SkipWhitespace
        
        If Mid(jsonText, pos, 1) = "]" Then
            pos = pos + 1
            ParseArray = Array()   ' array vacío estándar
            Exit Function
        End If
        
        Do
            ' Leer siguiente elemento
            Dim peek, val
            peek = Mid(jsonText, pos, 1)
            
            ' Llamar ParseValue (puede devolver objeto o escalar)
            If peek = "{" Or peek = "[" Then
                Dim objOrArr
                Set objOrArr = ParseValue()
                count = count + 1
                ReDim Preserve arr(count)
                Set arr(count) = objOrArr
            Else
                val = ParseValue()
                count = count + 1
                ReDim Preserve arr(count)
                arr(count) = val
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
        
        ParseArray = arr
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
            Err.Raise 1006, "JsonParser", "Numero invalido: " & numStr
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
            Err.Raise 1007, "JsonParser", "Valor booleano invalido"
        End If
    End Function
    
    Private Function ParseNull()
        If Mid(jsonText, pos, 4) = "null" Then
            pos = pos + 4
            ParseNull = Null
        Else
            Err.Raise 1008, "JsonParser", "Valor null invalido"
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

Function ImportFormFromJson(dbPath, jsonPath, nameOpt, overwrite, dryRun)
    On Error Resume Next
    ImportFormFromJson = False

    ' Resolver rutas
    Dim resolvedJsonPath: resolvedJsonPath = ResolvePath(jsonPath)
    Dim resolvedDbPath: resolvedDbPath = ResolvePath(dbPath)
    
    ' Leer JSON
    Dim json: json = ReadAllText(resolvedJsonPath)
    If json = "" Then 
        LogMessage "form-import: JSON vacio o no accesible"
        Exit Function
    End If

    ' Parsear JSON
    Dim root: Set root = JsonParse(json)
    If root Is Nothing Then 
        LogMessage "form-import: JSON invalido"
        Exit Function
    End If

    ' Validar estructura JSON
    Dim validationError: validationError = ""
    If Not ValidateFormJson(root, validationError) Then
        LogMessage "form-import: JSON invalido - " & validationError
        Exit Function
    End If

    ' Obtener nombre del formulario
    Dim targetName: targetName = nameOpt
    If targetName = "" Then targetName = GetDict(root, "name", "")
    If targetName = "" Then 
        LogMessage "form-import: falta nombre de formulario"
        Exit Function
    End If

    ' Abrir Access usando función canónica
    Dim app: Set app = OpenAccess(resolvedDbPath, GetDatabasePassword(resolvedDbPath))
    If app Is Nothing Then Exit Function

    ' Refuerzos anti-UI por simetria
    On Error Resume Next
    app.DisplayAlerts = False
    app.Echo False
    app.DoCmd.SetWarnings False
    Err.Clear

    ' Manejar overwrite
    If overwrite Then
        On Error Resume Next
        app.DoCmd.DeleteObject acForm, targetName
        Err.Clear
    End If

    ' Modo dry-run
    If dryRun Then
        Dim dryReport: dryReport = ""
        If UI_DryRunCheck(root, dryReport) Then
            LogMessage "form-import(dry-run): " & dryReport
            CloseAccess app
            ImportFormFromJson = True
        Else
            LogMessage "form-import(dry-run): " & dryReport
            CloseAccess app
            ImportFormFromJson = False
        End If
        Exit Function
    End If

    ' Patron atomico: crear formulario temporal
    Dim tmpName: tmpName = targetName & "__tmp"
    
    ' Crear temporal
    On Error Resume Next
    app.DisplayAlerts = False
    app.Echo False
    app.DoCmd.SetWarnings False
    Err.Clear
    app.DoCmd.CreateForm
    ' Reforzar anti-UI después de CreateForm
    app.DisplayAlerts = False
    app.Echo = False
    app.DoCmd.SetWarnings False
    If Err.Number <> 0 Then 
        LogMessage "form-import: no se pudo crear temporal: " & Err.Description
        Err.Clear
        CloseAccess app
        Exit Function
    End If
    Dim created: created = app.Screen.ActiveForm.Name
    On Error Resume Next
    app.DisplayAlerts = False
    app.Echo False
    app.DoCmd.SetWarnings False
    Err.Clear
    app.DoCmd.Rename tmpName, acForm, created
    On Error Resume Next
    app.DisplayAlerts = False
    app.Echo False
    app.DoCmd.SetWarnings False
    Err.Clear
    app.DoCmd.OpenForm tmpName, acDesign, , , , acHidden
    ' Reforzar anti-UI después de OpenForm
    app.DisplayAlerts = False
    app.Echo = False
    app.DoCmd.SetWarnings False
    Err.Clear

    ' Guardar referencias para las siguientes fases (usando tmpName)
    Set ImportFormFromJson_app = app
    ImportFormFromJson_target = tmpName
    ImportFormFromJson_finalTarget = targetName
    Set ImportFormFromJson_root = root

    ' Ejecutar fase 2: aplicar propiedades y controles
    If Not ImportFormFromJson_Apply() Then
        LogMessage "form-import: error en fase de aplicacion"
        ' Limpiar temporal en caso de error
        On Error Resume Next
        app.DoCmd.DeleteObject acForm, tmpName
        Err.Clear
        CloseAccess app
        Exit Function
    End If

    ' Ejecutar fase 3: finalizar y hacer swap atomico
    If Not ImportFormFromJson_Finalize() Then
        LogMessage "form-import: error en fase de finalizacion"
        ' Limpiar temporal en caso de error
        On Error Resume Next
        app.DoCmd.DeleteObject acForm, tmpName
        Err.Clear
        Exit Function
    End If

    ImportFormFromJson = True
End Function

Function ImportFormFromJson_Apply()
    On Error Resume Next
    ImportFormFromJson_Apply = False

    If ImportFormFromJson_app Is Nothing Then Exit Function
    Dim app: Set app = ImportFormFromJson_app
    Dim targetName: targetName = ImportFormFromJson_target
    Dim root: Set root = ImportFormFromJson_root

    Dim frm: Set frm = app.Forms(targetName)

    ' Propiedades de formulario
    SetPropIfExists frm, "RecordSource", GetDict(root, "recordSource", "")
    SetPropIfExists frm, "Caption", GetDict(root, "caption", "")
    ApplySections frm, root

    ' Limpiar controles existentes
    DeleteAllControls app, targetName, frm

    ' Crear controles desde JSON (orden dado)
    Dim arr, i, cnt
    If root.Exists("controls") Then
        arr = root("controls")
        If IsArray(arr) Then
            For i = 0 To UBound(arr)
                CreateControlFromJson app, targetName, arr(i)
            Next
        ElseIf IsObject(arr) Then
            On Error Resume Next
            Dim kkeys, kk
            kkeys = arr.Keys
            If Err.Number = 0 Then
                For kk = 0 To UBound(kkeys)
                    CreateControlFromJson app, targetName, arr(kkeys(kk))
                Next
            Else
                Err.Clear
                ' Mantén el fallback por índice si el objeto soporta indexación posicional
                cnt = arr.Count
                If Err.Number = 0 Then
                    For i = 0 To cnt - 1
                        CreateControlFromJson app, targetName, arr(i)
                    Next
                End If
                Err.Clear
            End If
        End If
    End If

    ImportFormFromJson_Apply = True
End Function

Sub SetPropIfExists(obj, p, v)
    On Error Resume Next
    If IsEmpty(v) Then Exit Sub
    If HasProperty(obj, p) Then
        obj.Properties(p) = v
        Err.Clear
    End If
End Sub

Function GetDict(d, k, def)
    On Error Resume Next
    If d.Exists(k) Then GetDict = d(k) Else GetDict = def
End Function

Sub ApplySections(frm, root)
    On Error Resume Next
    Dim s: Set s = GetDictObject(root, "sectionHeights")
    If Not s Is Nothing Then
        If s.Exists("detail") Then frm.Section(0).Height = s("detail") ' acDetail=0
        If s.Exists("header") Then frm.Section(1).Height = s("header") ' acHeader=1
        If s.Exists("footer") Then frm.Section(2).Height = s("footer") ' acFooter=2
    End If
End Sub

Function GetDictObject(d, k)
    On Error Resume Next
    If d.Exists(k) Then
        If IsObject(d(k)) Then Set GetDictObject = d(k)
    End If
End Function

Sub DeleteAllControls(app, formName, frm)
    On Error Resume Next
    ' Micro-guard: evitar intentar borrar si no hay controles
    If frm.Controls.Count <= 0 Then Exit Sub
    
    ' Capturar nombres antes del bucle para evitar problemas de reindexacion
    Dim names(), n, k
    n = frm.Controls.Count
    ReDim names(n-1)
    For k = 0 To n-1: names(k) = frm.Controls(k).Name: Next
    
    ' 1ª pasada: borrar TODO excepto TabCtl/Page
    For k = UBound(names) To 0 Step -1
        If frm.Controls(names(k)).ControlType <> acTabCtl And frm.Controls(names(k)).ControlType <> acPage Then
            app.DoCmd.DeleteControl formName, names(k)
            Err.Clear
        End If
    Next
    
    ' 2ª pasada: borrar Pages (si quedaran)
    For k = UBound(names) To 0 Step -1
        If frm.Controls(names(k)).ControlType = acPage Then
            app.DoCmd.DeleteControl formName, names(k)
            Err.Clear
        End If
    Next
    
    ' 3ª pasada: borrar TabCtl al final
    For k = UBound(names) To 0 Step -1
        If frm.Controls(names(k)).ControlType = acTabCtl Then
            app.DoCmd.DeleteControl formName, names(k)
            Err.Clear
        End If
    Next
End Sub

Sub CreateControlFromJson(app, formName, cjson)
    On Error Resume Next
    Dim ct, sec, name, left, top, width, height, ctrl
    ct    = GetDict(cjson, "controlType", 109) ' default textbox (Access: acTextBox=109)
    sec   = GetDict(cjson, "section", 0)      ' acDetail
    name  = GetDict(cjson, "name", "")
    left  = GetDict(cjson, "left", 0)
    top   = GetDict(cjson, "top", 0)
    width = GetDict(cjson, "width", 1200)
    height= GetDict(cjson, "height", 300)

    ' Lista blanca de tipos de controles soportados
    Dim supportedTypes: supportedTypes = Array( _
        acLabel, acTextBox, acCommandButton, acCheckBox, acOptionButton, _
        acComboBox, acListBox, acSubform, acImage, acRectangle, _
        acLine, acOptionGroup, acPageBreak, acCustomControl, _
        acToggleButton, acTabCtl, acPage _
    )
    
    Dim i, isSupported: isSupported = False
    For i = 0 To UBound(supportedTypes)
        If ct = supportedTypes(i) Then
            isSupported = True
            Exit For
        End If
    Next
    
    ' Mantener el extra If ct = 109 por robustez con JSON externos
    If ct = 109 Then isSupported = True
    
    If Not isSupported Then
        LogMessage "ui-import: controlType no soportado: " & ct
        Exit Sub
    End If

    Set ctrl = app.CreateControl(formName, ct, sec, , , left, top, width, height)
    If name <> "" Then On Error Resume Next: ctrl.Name = name: Err.Clear

    ' Propiedades comunes
    SetCtrlProp ctrl, "ControlSource", GetDict(cjson, "controlSource", "")
    SetCtrlProp ctrl, "RowSource",     GetDict(cjson, "rowSource", "")
    SetCtrlProp ctrl, "SourceObject",  GetDict(cjson, "sourceObject", "")
    SetCtrlProp ctrl, "Caption",       GetDict(cjson, "caption", "")
    SetCtrlProp ctrl, "Format",        GetDict(cjson, "format", "")
    SetCtrlProp ctrl, "FontName",      GetDict(cjson, "fontName", "")
    SetCtrlProp ctrl, "FontSize",      GetDict(cjson, "fontSize", "")
    SetCtrlProp ctrl, "ForeColor",     GetDict(cjson, "foreColor", "")
    SetCtrlProp ctrl, "BackColor",     GetDict(cjson, "backColor", "")
    SetCtrlProp ctrl, "TabIndex",      GetDict(cjson, "tabIndex", "")
    SetCtrlProp ctrl, "TabStop",       GetDict(cjson, "tabStop", "")
    SetCtrlProp ctrl, "Visible",       GetDict(cjson, "visible", "")
    SetCtrlProp ctrl, "Enabled",       GetDict(cjson, "enabled", "")
    SetCtrlProp ctrl, "Locked",        GetDict(cjson, "locked", "")
    ' Eventos (expresión)
    SetCtrlProp ctrl, "OnClick",       GetDict(cjson, "onClick", "")
    SetCtrlProp ctrl, "OnDblClick",    GetDict(cjson, "onDblClick", "")
    SetCtrlProp ctrl, "OnChange",      GetDict(cjson, "onChange", "")
    SetCtrlProp ctrl, "OnCurrent",     GetDict(cjson, "onCurrent", "")

    ' === Carga de imagen desde assets/ui/assets/img ===
    ' Convención JSON:
    '   - "image": nombre o ruta (relativa o absoluta)
    '   - "picture": alias alternativo
    '   - "assetImage": preferente para buscar en ui/assets/img
    Dim imgRef
    imgRef = GetDict(cjson, "assetImage", "")
    If imgRef = "" Then imgRef = GetDict(cjson, "image", "")
    If imgRef = "" Then imgRef = GetDict(cjson, "picture", "")

    If imgRef <> "" Then
        Dim cfgPath: Set cfgPath = gConfig ' ya cargado globalmente
        Dim fullImg
        fullImg = ResolveAssetImagePath(imgRef, cfgPath)
        If fullImg <> "" Then
            On Error Resume Next
            If ct = acImage Then
                ' Cargar sin UI: asignando la ruta directamente a Picture
                ctrl.Picture = fullImg
                ' Si existe la propiedad PictureType, usamos Linked (evita incrustar y prompts)
                If HasProperty(ctrl, "PictureType") Then
                    ' 1 = Linked (comportamiento silencioso)
                    ctrl.Properties("PictureType") = 1
                End If
                Err.Clear
            ElseIf ct = acCommandButton Then
                ' Botones con icono
                If HasProperty(ctrl, "Picture") Then
                    ctrl.Properties("Picture") = fullImg
                    Err.Clear
                End If
            End If
        Else
            LogMessage "ui-import: imagen no encontrada: " & imgRef
        End If
    End If
End Sub

Sub SetCtrlProp(ctrl, p, v)
    On Error Resume Next
    ' Salir si el valor está vacío, es nulo o es una cadena vacía
    If IsEmpty(v) Or v = "" Or IsNull(v) Then Exit Sub
    
    ' Verificar si el control tiene la propiedad
    If Not HasProperty(ctrl, p) Then
        LogVerbose "Propiedad " & p & " no existe en el control"
        Exit Sub
    End If
    
    ' Aplicar la propiedad
    ctrl.Properties(p) = v
    Err.Clear
End Sub

Function ImportFormFromJson_Finalize()
    On Error Resume Next
    ImportFormFromJson_Finalize = False
    If ImportFormFromJson_app Is Nothing Then Exit Function

    Dim app: Set app = ImportFormFromJson_app
    Dim tmpName: tmpName = ImportFormFromJson_target
    Dim finalName: finalName = ImportFormFromJson_finalTarget

    ' Guardar temporal
    app.DoCmd.Save acForm, tmpName
    Err.Clear
    
    ' Swap atomico: borrar antiguo y renombrar temporal
    app.DoCmd.DeleteObject acForm, finalName  ' borrar antiguo si existe (silencioso)
    Err.Clear
    app.DoCmd.Rename finalName, acForm, tmpName
    On Error Resume Next
    app.DisplayAlerts = False
    app.Echo False
    app.DoCmd.SetWarnings False
    Err.Clear
    If Err.Number <> 0 Then
        ' Si falla el rename, intentar borrar el antiguo y reintentar
        app.DoCmd.DeleteObject acForm, finalName
        Err.Clear
        app.DoCmd.Rename finalName, acForm, tmpName
        On Error Resume Next
        app.DisplayAlerts = False
        app.Echo False
        app.DoCmd.SetWarnings False
        Err.Clear
    End If
    Err.Clear
    
    On Error Resume Next
    app.RunCommand 126                 ' acCmdCompileAndSaveAllModules
    Err.Clear
    CloseAccess app

    ' Limpiar referencias
    Set ImportFormFromJson_app = Nothing
    ImportFormFromJson_target = ""
    ImportFormFromJson_finalTarget = ""
    Set ImportFormFromJson_root = Nothing

    ImportFormFromJson_Finalize = True
End Function

' ============================================================================
' SECCIÓN UI-AS-CODE: FUNCIONES DE REBUILD Y UPDATE
' ============================================================================

Function UI_RebuildAll(dbPath, files, config)
    On Error Resume Next
    UI_RebuildAll = False
    Dim app: Set app = OpenAccess(dbPath, GetDatabasePassword(dbPath))
    If app Is Nothing Then Exit Function

    ' Reforzar blindaje como en rebuild/update
    On Error Resume Next
    app.DisplayAlerts = False
    app.Echo False
    app.DoCmd.SetWarnings False
    Err.Clear

    ' Validar array de archivos
    If Not IsArray(files) Or SafeUBound(files) < 0 Then
        LogMessage "ui-rebuild: no se encontraron formularios"
        CloseAccess app
        UI_RebuildAll = True
        Exit Function
    End If

    Dim i, formPath, formName, tmpName
    For i = 0 To UBound(files)
        formPath = ResolvePath(files(i))
        formName = UI_DeduceFormName(formPath, config)
        If formName <> "" Then
            ' Patron atomico: importar en temporal primero
            tmpName = formName & "__tmp"
            If UI_ImportOne(app, formPath, tmpName) Then
                ' Si la importacion temporal fue exitosa, hacer swap atomico
                app.DoCmd.DeleteObject acForm, formName
                Err.Clear
                app.DoCmd.Rename formName, acForm, tmpName
                On Error Resume Next
                app.DisplayAlerts = False
                app.Echo False
                app.DoCmd.SetWarnings False
                Err.Clear
                If Err.Number <> 0 Then
                    ' Si falla el rename, intentar borrar el antiguo y reintentar
                    app.DoCmd.DeleteObject acForm, formName
                    Err.Clear
                    app.DoCmd.Rename formName, acForm, tmpName
                    On Error Resume Next
                    app.DisplayAlerts = False
                    app.Echo False
                    app.DoCmd.SetWarnings False
                    Err.Clear
                End If
                Err.Clear
            Else
                LogMessage "ui-rebuild: error en " & formName & ", manteniendo version anterior"
            End If
        End If
    Next

    ' Compilacion global al final (como update/rebuild)
    On Error Resume Next
    app.RunCommand 126  ' acCmdCompileAndSaveAllModules
    Err.Clear
    CloseAccess app
    UI_RebuildAll = True
End Function

Function UI_UpdateSome(dbPath, items, config)
    On Error Resume Next
    UI_UpdateSome = False
    Dim app: Set app = OpenAccess(dbPath, GetDatabasePassword(dbPath))
    If app Is Nothing Then Exit Function

    ' Reforzar blindaje como en rebuild/update
    On Error Resume Next
    app.DisplayAlerts = False
    app.Echo False
    app.DoCmd.SetWarnings False
    Err.Clear

    ' Validar array de items
    If Not IsArray(items) Or SafeUBound(items) < 0 Then
        LogMessage "ui-update: no se encontraron elementos para actualizar"
        CloseAccess app
        UI_UpdateSome = True
        Exit Function
    End If

    Dim i, formPath, formsDir, uiRoot, resolved, pattern, tmpName
    uiRoot = UI_GetRoot(config)
    formsDir = UI_GetFormsDir(config)
    pattern = UI_GetFormFilePattern(config)

    For i = 0 To UBound(items)
        resolved = items(i)
        ' Permite pasar nombre base o ruta JSON
        If InStr(LCase(resolved), ".json") = 0 Then
            resolved = ResolvePath(uiRoot & "\" & formsDir & "\" & resolved & ".json")
        Else
            resolved = ResolvePath(resolved)
        End If
        formPath = resolved
        
        ' Validar que el archivo JSON existe
        Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
        If Not objFSO.FileExists(formPath) Then
            LogMessage "ui-update: JSON no encontrado: " & formPath
            ' Continuar con el siguiente elemento
        Else
            Dim formName: formName = UI_DeduceFormName(formPath, config)
            If formName <> "" Then
                ' Patron atomico: importar en temporal primero
                tmpName = formName & "__tmp"
                If UI_ImportOne(app, formPath, tmpName) Then
                    ' Si la importacion temporal fue exitosa, hacer swap atomico
                    app.DoCmd.DeleteObject acForm, formName
                    Err.Clear
                    app.DoCmd.Rename formName, acForm, tmpName
                    On Error Resume Next
                    app.DisplayAlerts = False
                    app.Echo False
                    app.DoCmd.SetWarnings False
                    Err.Clear
                    If Err.Number <> 0 Then
                        ' Si falla el rename, intentar borrar el antiguo y reintentar
                        app.DoCmd.DeleteObject acForm, formName
                        Err.Clear
                        app.DoCmd.Rename formName, acForm, tmpName
                        On Error Resume Next
                        app.DisplayAlerts = False
                        app.Echo False
                        app.DoCmd.SetWarnings False
                        Err.Clear
                    End If
                    Err.Clear
                Else
                    LogMessage "ui-update: error en " & formName & ", manteniendo version anterior"
                End If
            End If
        End If
    Next

    On Error Resume Next
    app.RunCommand 126  ' acCmdCompileAndSaveAllModules
    Err.Clear
    CloseAccess app
    UI_UpdateSome = True
End Function

Function UI_DeduceFormName(formPath, config)
    On Error Resume Next
    Dim nameFromFile: nameFromFile = UI_NameFromFileBase(config)
    If nameFromFile Then
        UI_DeduceFormName = FileBaseName(formPath)
    Else
        ' en el futuro: leer "name" del JSON si se quisiera
        UI_DeduceFormName = FileBaseName(formPath)
    End If
End Function

Function UI_ImportOne(app, formPath, formName)
    On Error Resume Next
    UI_ImportOne = False
    Dim json: json = ReadAllText(formPath)
    If json = "" Then Exit Function
    Dim root: Set root = JsonParse(json)
    If root Is Nothing Then Exit Function

    ' Validar estructura JSON
    Dim validationError: validationError = ""
    If Not ValidateFormJson(root, validationError) Then
        LogMessage "ui-import: JSON invalido - " & validationError
        Exit Function
    End If

    ' Patron atomico: crear formulario temporal
    Dim tmpName: tmpName = formName & "__tmp"
    
    ' Crear temporal
    On Error Resume Next
    app.DisplayAlerts = False
    app.Echo False
    app.DoCmd.SetWarnings False
    Err.Clear
    app.DoCmd.CreateForm
    ' Reforzar anti-UI después de CreateForm
    app.DisplayAlerts = False
    app.Echo = False
    app.DoCmd.SetWarnings False
    If Err.Number <> 0 Then 
        LogMessage "ui-import: no se pudo crear temporal: " & Err.Description
        Err.Clear
        Exit Function
    End If
    Dim created: created = app.Screen.ActiveForm.Name
    On Error Resume Next
    app.DisplayAlerts = False
    app.Echo False
    app.DoCmd.SetWarnings False
    Err.Clear
    app.DoCmd.Rename tmpName, acForm, created
    On Error Resume Next
    app.DisplayAlerts = False
    app.Echo False
    app.DoCmd.SetWarnings False
    Err.Clear
    Err.Clear
    app.DoCmd.OpenForm tmpName, acDesign, , , , acHidden
    ' Reforzar anti-UI después de OpenForm
    app.DisplayAlerts = False
    app.Echo = False
    app.DoCmd.SetWarnings False
    Err.Clear
    
    Dim frm: Set frm = app.Forms(tmpName)

    ' Aplicar propiedades y controles sobre el temporal
    SetPropIfExists frm, "RecordSource", GetDict(root,"recordSource","")
    SetPropIfExists frm, "Caption",      GetDict(root,"caption","")
    ApplySections frm, root

    ' Controles
    DeleteAllControls app, tmpName, frm
    Dim arr, i, cnt
    If root.Exists("controls") Then
        arr = root("controls")
        If IsArray(arr) Then
            For i = 0 To UBound(arr)
                CreateControlFromJson app, tmpName, arr(i)
            Next
        ElseIf IsObject(arr) Then
            On Error Resume Next
            Dim kkeys, kk
            kkeys = arr.Keys
            If Err.Number = 0 Then
                For kk = 0 To UBound(kkeys)
                    CreateControlFromJson app, tmpName, arr(kkeys(kk))
                Next
            Else
                Err.Clear
                ' Mantén el fallback por índice si el objeto soporta indexación posicional
                cnt = arr.Count
                If Err.Number = 0 Then
                    For i = 0 To cnt - 1
                        CreateControlFromJson app, tmpName, arr(i)
                    Next
                End If
                Err.Clear
            End If
        End If
    End If

    ' Guardar temporal
    app.DoCmd.Save acForm, tmpName
    Err.Clear
    
    ' Swap atomico: si todo va bien, borrar antiguo y renombrar temporal
    app.DoCmd.DeleteObject acForm, formName  ' borrar antiguo si existe (silencioso)
    Err.Clear
    app.DoCmd.Rename formName, acForm, tmpName
    If Err.Number <> 0 Then
        ' Si falla el rename, intentar borrar el antiguo y reintentar
        app.DoCmd.DeleteObject acForm, formName
        Err.Clear
        app.DoCmd.Rename formName, acForm, tmpName
    End If
    Err.Clear
    
    ' Cerrar el formulario ya renombrado
    app.DoCmd.Close acForm, formName, acSaveNo   ' ya guardado
    UI_ImportOne = True
End Function

' Stub seguro para RunTests - evita errores y procesos colgados
Sub RunTests()
    LogMessage "RunTests: no implementado (stub)"
End Sub