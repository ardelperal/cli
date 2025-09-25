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

' Constantes de objetos de Access (tipos de objeto)
Const acTable = 0
Const acQuery = 1
Const acForm = 2
Const acReport = 3
Const acMacro = 4
Const acModule = 5
Const acClassModule = 100

' Constantes de vistas de Access
Const acViewNormal = 0
Const acViewDesign = 1
Const acViewPreview = 2

' Constantes de modo de ventana
Const acHidden = 1
Const acNormal = 0

' Constantes de guardado
Const acSaveNo = 0
Const acSaveYes = 1

Const acDefault = -1

' Constantes de comandos de Access
Const acCmdCompileAndSaveAllModules = 126

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

' ============================================================================
' SECCIÓN 2: VARIABLES GLOBALES
' ============================================================================

Dim objFSO, objArgs, objAccess, objConfig
Dim gVerbose, gQuiet, gDryRun, gDebug
Dim gDbPath, gPassword, gOutputPath, gConfigPath, gScriptPath, gScriptDir
Dim g_ModulesSrcPath, g_ModulesExtensions, g_ModulesIncludeSubdirs
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
        LogMessage "Archivo de configuracion no encontrado: " & configPath & ". Usando valores por defecto."
        Set LoadConfig = config
        Exit Function
    Else
        LogMessage "Archivo de configuracion encontrado: " & configPath
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
                        
                        LogMessage "Procesando: [" & section & "] " & Trim(parts(0)) & " = " & value & " -> clave: " & key
                        
                        ' Resolver rutas relativas para ciertos valores (excepto patrones y booleanos)
                        If (InStr(UCase(key), "PATH") > 0 Or InStr(UCase(key), "FILE") > 0) And InStr(UCase(key), "PATTERN") = 0 And InStr(UCase(key), "FROMFILEBASE") = 0 Then
                            If value <> "" And fso.GetAbsolutePathName(value) <> value Then
                                value = gScriptDir & "\" & value
                                LogMessage "Ruta resuelta: " & value
                            End If
                        End If
                        
                        If config.Exists(key) Then
                            config(key) = value
                            LogMessage "Clave actualizada: " & key & " = " & value
                        Else
                            config.Add key, value
                            LogMessage "Clave agregada: " & key & " = " & value
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

Function EnsureConfigLoaded()
    On Error Resume Next
    If (gConfig Is Nothing) Then
        Set gConfig = LoadConfig(gConfigPath)
    End If
    EnsureConfigLoaded = Not (gConfig Is Nothing)
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

Function ReadAllText(path)
    On Error Resume Next
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = CInt(CfgGet(gConfig, "VBA_StreamTypeText", "VBA.StreamTypeText", "2")) ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile path
    ReadAllText = stream.ReadText
    stream.Close
    If Err.Number <> 0 Then ReadAllText = ""
    Err.Clear
End Function

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

Function ConcatArrays(arr1, arr2)
    On Error Resume Next
    Dim result, i, totalSize
    
    ' Handle empty arrays
    If Not IsArray(arr1) And Not IsArray(arr2) Then
        ConcatArrays = Array()
        Exit Function
    End If
    
    If Not IsArray(arr1) Then
        ConcatArrays = arr2
        Exit Function
    End If
    
    If Not IsArray(arr2) Then
        ConcatArrays = arr1
        Exit Function
    End If
    
    ' Calculate total size
    totalSize = UBound(arr1) + UBound(arr2) + 2
    ReDim result(totalSize - 1)
    
    ' Copy first array
    For i = 0 To UBound(arr1)
        result(i) = arr1(i)
    Next
    
    ' Copy second array
    For i = 0 To UBound(arr2)
        result(UBound(arr1) + 1 + i) = arr2(i)
    Next
    
    ConcatArrays = result
End Function

Function GetExtension(filePath)
    On Error Resume Next
    Dim fso: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ext: ext = fso.GetExtensionName(filePath)
    If ext <> "" Then
        GetExtension = "." & ext
    Else
        GetExtension = ""
    End If
End Function

Function JoinTailArgs(args)
    On Error Resume Next
    Dim result, i
    result = ""
    
    If IsArray(args) And UBound(args) >= 1 Then
        For i = 1 To UBound(args)
            If result <> "" Then result = result & " "
            result = result & args(i)
        Next
    End If
    JoinTailArgs = result
End Function

' ============================================================================
' SECCIÓN 3.1: FUNCIONES DE CONFIGURACIÓN UI AS CODE
' ============================================================================

Function UI_GetRoot(config)
    On Error Resume Next
    If (config Is Nothing) Then Set config = gConfig
    If (config Is Nothing) Then UI_GetRoot = ResolvePath(".\ui") : Exit Function
    Dim rootValue: rootValue = CfgGet(config,"UI_Root","UI.Root",".\ui")
    ' Si ya es una ruta absoluta, no la resolvemos de nuevo
    If objFSO.GetAbsolutePathName(rootValue) = rootValue Then
        UI_GetRoot = rootValue
    Else
        UI_GetRoot = ResolvePath(rootValue)
    End If
End Function

Function UI_GetFormsDir(config)
    On Error Resume Next
    If (config Is Nothing) Then Set config = gConfig
    If (config Is Nothing) Then UI_GetFormsDir = "forms" : Exit Function
    UI_GetFormsDir = CfgGet(config,"UI_FormsDir","UI.FormsDir","forms")
End Function

Function UI_GetAssetsDir(config)
    On Error Resume Next
    If (config Is Nothing) Then Set config = gConfig
    If (config Is Nothing) Then UI_GetAssetsDir = "assets" : Exit Function
    UI_GetAssetsDir = CfgGet(config,"UI_AssetsDir","UI.AssetsDir","assets")
End Function

Function UI_GetIncludeSubdirs(config)
    On Error Resume Next
    If (config Is Nothing) Then Set config = gConfig
    If (config Is Nothing) Then UI_GetIncludeSubdirs = True : Exit Function
    UI_GetIncludeSubdirs = ToBool(CfgGet(config,"UI_IncludeSubdirectories","UI.IncludeSubdirectories","true"),True)
End Function

Function UI_GetFormFilePattern(config)
    On Error Resume Next
    If (config Is Nothing) Then Set config = gConfig
    If (config Is Nothing) Then UI_GetFormFilePattern = "*.json" : Exit Function
    UI_GetFormFilePattern = CfgGet(config,"UI_FormFilePattern","UI.FormFilePattern","*.json")
End Function

Function UI_NameFromFileBase(config)
    On Error Resume Next
    If (config Is Nothing) Then Set config = gConfig
    If (config Is Nothing) Then UI_NameFromFileBase = True : Exit Function
    UI_NameFromFileBase = ToBool(CfgGet(config,"UI_NameFromFileBase","UI.NameFromFileBase","true"),True)
End Function

Function UI_GetAssetsImgDir(config)
    On Error Resume Next
    If (config Is Nothing) Then Set config = gConfig
    If (config Is Nothing) Then UI_GetAssetsImgDir = "img" : Exit Function
    UI_GetAssetsImgDir = CfgGet(config,"UI_AssetsImgDir","UI.AssetsImgDir","img")
End Function

Function UI_GetAssetsImgExtensions(config)
    On Error Resume Next
    If (config Is Nothing) Then Set config = gConfig
    If (config Is Nothing) Then UI_GetAssetsImgExtensions = ".png,.jpg,.jpeg,.gif,.bmp" : Exit Function
    Dim exts: exts = CfgGet(config,"UI_AssetsImgExtensions","UI.AssetsImgExtensions",".png,.jpg,.jpeg,.gif,.bmp")
    UI_GetAssetsImgExtensions = LCase(exts)
End Function

Function UI_StrictProperties(config)
    On Error Resume Next
    If (config Is Nothing) Then Set config = gConfig
    If (config Is Nothing) Then UI_StrictProperties = False : Exit Function
    UI_StrictProperties = ToBool(CfgGet(config,"UI_StrictProperties","UI.StrictProperties","false"), False)
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
    WScript.Echo "  extract-modules [db_path]"
    WScript.Echo "    Extrae modulos VBA hacia archivos fuente"
    WScript.Echo ""
    WScript.Echo "  list-objects <db_path> [--password <pwd>]"
    WScript.Echo "    Lista objetos de la base de datos"
    WScript.Echo ""
    WScript.Echo "  rebuild"
    WScript.Echo "    Reconstruye modulos VBA desde los archivos fuente configurados"
    WScript.Echo "    - Usa MODULES_SrcPath del archivo de configuracion"
    WScript.Echo "    - Formatos soportados: .bas (modulos), .cls (clases)"
    WScript.Echo "    - Usa DATABASE_DefaultPath del archivo de configuracion"
    WScript.Echo "    - Requiere activar ""Trust access to the VBA project object model"" en Access"
    WScript.Echo ""
    WScript.Echo "  update <modulo1,modulo2,...>"
    WScript.Echo "    Actualiza modulos VBA especificos usando los fuentes locales"
    WScript.Echo ""
    WScript.Echo "  export-form <db_path> <form_name> [--output <path>] [--password <pwd>]"
    WScript.Echo "    Exporta un formulario a JSON"
    WScript.Echo ""
    WScript.Echo "  ui-rebuild [db_path]"
    WScript.Echo "    Reconstruye todos los formularios JSON ubicados en /ui hacia Access"
    WScript.Echo ""
    WScript.Echo "  ui-update [db_path] <form1.json> [form2.json ...]"
    WScript.Echo "    Actualiza formularios especificos desde archivos JSON"
    WScript.Echo ""
    WScript.Echo "  ui-touch <formName|form.json>"
    WScript.Echo "    Atajo para actualizar un unico formulario"
    WScript.Echo ""
    WScript.Echo "  test"
    WScript.Echo "    Ejecuta la bateria interna de pruebas"
    WScript.Echo ""
    WScript.Echo "OPCIONES:"
    WScript.Echo "  --config <path>       - Archivo de configuracion (por defecto: cli.ini)"
    WScript.Echo "  --password <pwd>      - Contrasena de la base de datos"
    WScript.Echo "  --verbose             - Salida detallada"
    WScript.Echo "  --quiet               - Salida minima"
    WScript.Echo "  --debug               - Activa diagnosticos adicionales"
    WScript.Echo "  --help                - Muestra esta ayuda"
    WScript.Echo ""
    WScript.Echo "MODIFICADORES:"
    WScript.Echo "  /dry-run | --dry-run  - Simula la ejecucion sin cambios"
    WScript.Echo "  /validate             - Valida la configuracion actual"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS:"
    WScript.Echo "  cscript cli.vbs extract-modules ""C:\mi_base.accdb"""
    WScript.Echo "  cscript cli.vbs list-objects ""C:\mi_base.accdb"" --verbose"
    WScript.Echo "  cscript cli.vbs rebuild"
    WScript.Echo "  cscript cli.vbs update ModuloA,ModuloB"
    WScript.Echo "  cscript cli.vbs export-form ""C:\mi_base.accdb"" MainForm --output ""C:\salida\MainForm.json"""
    WScript.Echo "  cscript cli.vbs ui-rebuild"
    WScript.Echo "  cscript cli.vbs ui-update ""C:\mi_base.accdb"" forms\MainForm.json forms\Users.json"
    WScript.Echo "  cscript cli.vbs ui-touch MainForm"
    WScript.Echo "  cscript cli.vbs test"
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
    Dim objApp
    
    LogVerbose "Abriendo Access: " & dbPath
    If password <> "" Then
        LogVerbose "Con password: (oculta)"
    End If
    
    On Error Resume Next
    Set objApp = CreateObject("Access.Application")
    
    If Err.Number <> 0 Then
        LogMessage "ERROR: No se pudo crear instancia de Access: " & Err.Description
        Set OpenAccess = Nothing
        Exit Function
    End If
    
    ' Configurar Access para modo silencioso (solo lo que NO requiere BD abierta)
    objApp.Visible = False
    objApp.UserControl = False
    objApp.AutomationSecurity = 1  ' msoAutomationSecurityForceDisable - IGUAL que rebuild
    
    ' Abrir la base de datos
    If password <> "" Then
        objApp.OpenCurrentDatabase dbPath, False, password
    Else
        objApp.OpenCurrentDatabase dbPath, False
    End If

    If Err.Number <> 0 Then
        LogMessage "ERROR: No se pudo abrir la base de datos: " & Err.Description
        objApp.Quit
        Set objApp = Nothing
        Set OpenAccess = Nothing
        Exit Function
    End If
    
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
    If Not gQuiet Then
        ' Reemplazar caracteres especiales para evitar problemas de codificación
        message = Replace(message, "ó", "o")
        message = Replace(message, "ñ", "n")
        message = Replace(message, "á", "a")
        message = Replace(message, "é", "e")
        message = Replace(message, "í", "i")
        message = Replace(message, "ú", "u")
        message = Replace(message, "Ó", "O")
        message = Replace(message, "Ñ", "N")
        message = Replace(message, "Á", "A")
        message = Replace(message, "É", "E")
        message = Replace(message, "Í", "I")
        message = Replace(message, "Ú", "U")
        WScript.Echo "[" & Now & "] " & message
    End If
End Sub

Sub LogVerbose(message)
    If gVerbose And Not gQuiet Then
        WScript.Echo "[VERBOSE] " & message
    End If
End Sub

Sub LogError(message)
    WScript.Echo "[ERROR] " & message
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

Function ExtractFormControls(db, formName)
    Dim controlsDict, app, frm
    Set controlsDict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    
    ' Abrir Access y el formulario
    Set app = OpenAccess(db.Name, gPassword)
    
    app.DoCmd.OpenForm formName, acDesign, , , , acHidden
    
    If Err.Number <> 0 Then
        LogMessage "Error: Formulario '" & formName & "' no encontrado o no se puede abrir", "ERROR"
        CloseAccess app
        Set ExtractFormControls = controlsDict
        Exit Function
    End If
    
    Set frm = app.Forms(formName)
    Set controlsDict = ExtractFormControlsInternal(frm)
    
    app.DoCmd.Close acForm, formName, acSaveNo
    CloseAccess app
    
    Set ExtractFormControls = controlsDict
End Function

Function ExtractFormControlsInternal(frm)
    Dim controlsDict, controlDict, ctl
    Set controlsDict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    
    For Each ctl In frm.Controls
        Set controlDict = CreateObject("Scripting.Dictionary")
        
        ' Propiedades básicas del control
        controlDict("name") = ctl.Name
        controlDict("controlType") = GetControlTypeName(ctl.ControlType)
        controlDict("left") = ctl.Left
        controlDict("top") = ctl.Top
        controlDict("width") = ctl.Width
        controlDict("height") = ctl.Height
        controlDict("visible") = ctl.Visible
        controlDict("enabled") = ctl.Enabled
        
        ' Propiedades específicas según el tipo de control
        Select Case ctl.ControlType
            Case acTextBox, acComboBox, acListBox
                On Error Resume Next
                controlDict("controlSource") = ctl.ControlSource
                controlDict("rowSource") = ctl.RowSource
                controlDict("rowSourceType") = ctl.RowSourceType
                controlDict("boundColumn") = ctl.BoundColumn
                controlDict("columnCount") = ctl.ColumnCount
                On Error GoTo 0
                
            Case acLabel
                On Error Resume Next
                controlDict("caption") = ctl.Caption
                On Error GoTo 0
                
            Case acCommandButton
                On Error Resume Next
                controlDict("caption") = ctl.Caption
                controlDict("onClick") = ctl.OnClick
                On Error GoTo 0
                
            Case acCheckBox, acOptionButton, acToggleButton
                On Error Resume Next
                controlDict("controlSource") = ctl.ControlSource
                controlDict("defaultValue") = ctl.DefaultValue
                On Error GoTo 0
                
            Case acSubform
                On Error Resume Next
                controlDict("sourceObject") = ctl.SourceObject
                controlDict("linkChildFields") = ctl.LinkChildFields
                controlDict("linkMasterFields") = ctl.LinkMasterFields
                On Error GoTo 0
        End Select
        
        ' Propiedades de formato
        On Error Resume Next
        controlDict("backColor") = ctl.BackColor
        controlDict("foreColor") = ctl.ForeColor
        controlDict("borderStyle") = ctl.BorderStyle
        controlDict("fontSize") = ctl.FontSize
        controlDict("fontName") = ctl.FontName
        controlDict("fontBold") = ctl.FontBold
        controlDict("fontItalic") = ctl.FontItalic
        On Error GoTo 0
        
        Set controlsDict(ctl.Name) = controlDict
    Next
    
    Set ExtractFormControlsInternal = controlsDict
End Function

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
    Dim jsonText, outputFile
    
    jsonText = ConvertToJSON(data)
    
    Set outputFile = objFSO.CreateTextFile(filePath, True, False)
    outputFile.Write jsonText
    outputFile.Close
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
Function CopyAsAnsi(srcPath, dstPath)
    CopyAsAnsi = False
    Dim fso, srcFile, dstFile, content
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    
    ' Leer archivo fuente usando TextStream con formato Unicode
    Set srcFile = fso.OpenTextFile(srcPath, 1, False, -1) ' ForReading, Unicode
    If Err.Number <> 0 Then
        Err.Clear
        ' Fallback: intentar como ASCII
        Set srcFile = fso.OpenTextFile(srcPath, 1, False, 0) ' ForReading, ASCII
        If Err.Number <> 0 Then
            Err.Clear
            Exit Function
        End If
    End If
    
    ' Leer todo el contenido
    content = srcFile.ReadAll()
    srcFile.Close
    
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    
    ' Para archivos .bas, asegurar que Option Explicit esté presente
    If LCase(Right(srcPath, 4)) = ".bas" Then
        content = EnsureOptionExplicit(content)
    End If
    
    ' Escribir archivo destino como ASCII/ANSI
    Set dstFile = fso.CreateTextFile(dstPath, True, False) ' Overwrite, ASCII
    dstFile.Write content
    dstFile.Close
    
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    
    On Error GoTo 0
    
    If gDebug Then LogMessage "[DEBUG] Archivo copiado como ANSI: " & srcPath & " -> " & dstPath
    CopyAsAnsi = True
End Function

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

' Función para escribir archivos de texto usando ADODB.Stream
Sub WriteTextFile(filePath, content)
    On Error Resume Next
    
    ' Crear directorio padre si no existe
    Dim parentDir
    parentDir = objFSO.GetParentFolderName(filePath)
    If Not objFSO.FolderExists(parentDir) Then
        CreateFolderRecursive parentDir
    End If
    
    ' Escribir archivo usando ADODB.Stream para preservar caracteres especiales
    Dim objStream
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "UTF-8"
    objStream.Open
    objStream.WriteText content, 0 ' adWriteChar
    objStream.SaveToFile filePath, 2 ' adSaveCreateOverWrite
    objStream.Close
    Set objStream = Nothing
    
    If Err.Number <> 0 Then
        LogMessage "Error escribiendo archivo: " & filePath & " - " & Err.Description
        Err.Clear
    End If
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
    
    If Err.Number = 0 Then
        ExportAccessObject = True
    Else
        LogMessage "Error exportando " & objectName & ": " & Err.Description
        Err.Clear
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
            RunTests
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
        ElseIf LCase(objArgs(i)) = "--password" Then
            If i < objArgs.Count - 1 Then
                gPassword = objArgs(i + 1)
                LogMessage "Password especificado via parametro"
                i = i + 1 ' Saltar el siguiente argumento (el password)
            Else
                WScript.Echo "Error: --password requiere un valor"
                WScript.Quit 1
            End If
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
                Dim config
                Set config = LoadConfig(gConfigPath)
                gDbPath = ResolvePath(config("DATABASE_DefaultPath"))
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
            Set config = LoadConfig(gConfigPath)
            gDbPath = ResolvePath(config("DATABASE_DefaultPath"))
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
            Set config = LoadConfig(gConfigPath)
            gDbPath = ResolvePath(config("DATABASE_DefaultPath"))
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
                
                ' Procesar argumentos opcionales para list-objects
                Dim i2, showSchema, outputToFile
                showSchema = False
                outputToFile = False
                
                For i2 = 2 To cleanArgCount - 1
                    If LCase(cleanArgs(i2)) = "--password" And i2 < cleanArgCount - 1 Then
                        gPassword = cleanArgs(i2 + 1)
                    ElseIf LCase(cleanArgs(i2)) = "--schema" Then
                        showSchema = True
                    ElseIf LCase(cleanArgs(i2)) = "--output" Then
                        outputToFile = True
                    End If
                Next
                
                If Not gDryRun Then
                    ListObjects gDbPath, showSchema, outputToFile
                Else
                    LogMessage "SIMULACION: Listaria objetos de " & gDbPath
                End If
            Else
                LogError "Faltan argumentos para list-objects"
                WScript.Echo "Uso: cscript cli.vbs list-objects <db_path> [--password <pwd>] [--schema] [--output]"
                WScript.Echo "  --password <pwd> : Especifica la contraseña de la base de datos"
                WScript.Echo "  --schema         : Muestra detalles de campos en las tablas"
                WScript.Echo "  --output         : Exporta resultados a archivo [nombre_bd]_listobjects.txt"
                ShowHelp
            End If
            
        Case "export-form"
            If cleanArgCount >= 3 Then
                Dim strDbPath2, strFormName2, strOutputPath2, strPassword2
                Dim i3
                
                ' Asignar argumentos básicos
                strDbPath2 = ResolvePath(cleanArgs(1))
                strFormName2 = cleanArgs(2)
                strOutputPath2 = ""
                strPassword2 = gPassword
                
                ' Procesar argumentos opcionales
                For i3 = 3 To cleanArgCount - 1
                    If LCase(cleanArgs(i3)) = "--output" And i3 < cleanArgCount - 1 Then
                        strOutputPath2 = cleanArgs(i3 + 1)
                    ElseIf LCase(cleanArgs(i3)) = "--password" And i3 < cleanArgCount - 1 Then
                        strPassword2 = cleanArgs(i3 + 1)
                    End If
                Next
                
                If Not gDryRun Then
                    ExportFormToJSON strDbPath2, strFormName2, strOutputPath2, strPassword2
                Else
                    LogMessage "SIMULACION: Exportaria formulario " & strFormName2 & " de " & strDbPath2
                End If
            Else
                LogError "Faltan argumentos para export-form"
                WScript.Echo "Uso: cscript cli.vbs export-form <db_path> <form_name> [--output <path>] [--password <pwd>]"
                WScript.Echo "  --output <path>  : Especifica la ruta de salida del archivo JSON"
                WScript.Echo "  --password <pwd> : Especifica la contraseña de la base de datos"
                ShowHelp
            End If
            
        Case "test"
            RunTests
            
        Case "ui-rebuild"
            ' ui-rebuild [db_path]
            Dim config4, dbArg4, uiRoot, formsDir, includeSub, pattern, formsPath, files, i4
            Set config4 = LoadConfig(gConfigPath)
            Call EnsureConfigLoaded()
            dbArg4 = ""
            If cleanArgCount > 1 Then dbArg4 = cleanArgs(1)
            If dbArg4 = "" Then dbArg4 = CfgGet(config4,"DATABASE_DefaultPath","DATABASE.DefaultPath","")
            gDbPath = ResolvePath(dbArg4)
            uiRoot = UI_GetRoot(config4)
            formsDir = UI_GetFormsDir(config4)
            includeSub = UI_GetIncludeSubdirs(config4)
            pattern = UI_GetFormFilePattern(config4)
            formsPath = uiRoot & "\" & formsDir
            LogMessage "Buscando archivos JSON en: " & formsPath
            LogMessage "Patron: " & pattern & ", IncludeSubdirs: " & includeSub
            files = EnumerateFiles(formsPath, pattern, includeSub)
            LogMessage "Archivos encontrados: " & (UBound(files) + 1)
            If Not UI_RebuildAll(gDbPath, files, config4) Then WScript.Quit 1 Else WScript.Quit 0
            
        Case "ui-update"
            ' ui-update [db_path] <form1.json> [form2.json ...]  (acepta rutas relativas o nombres base)
            Dim dbArg5, idx, list(), config5
            ReDim list(-1)
            Set config5 = LoadConfig(gConfigPath)
            Call EnsureConfigLoaded()
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
            Call EnsureConfigLoaded()
            dbArg6 = CfgGet(config6,"DATABASE_DefaultPath","DATABASE.DefaultPath","")
            If cleanArgCount < 2 Then WScript.Echo CleanTerminalText("uso: ui-touch <formName|form.json>"): WScript.Quit 1
            item = cleanArgs(1)
            Dim arr(0): arr(0) = item
            If Not UI_UpdateSome(ResolvePath(dbArg6), arr, config6) Then WScript.Quit 1 Else WScript.Quit 0
            
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
    
    WScript.Echo "=== ACTUALIZACION DE MODULOS ESPECIFICOS ==="
    
    ' Normalizar lista de módulos antes de abrir Access
    Dim list, i, name
    list = NormalizeModuleList(modulesArg)
    
    If IsEmpty(list) Or UBound(list) < 0 Then
        LogMessage "update: sin módulos para actualizar"
        UpdateModules = True
        Exit Function
    End If
    
    WScript.Echo "Modulos a actualizar: " & Join(list, ", ")
    
    ' PASO 1: Cerrar procesos de Access existentes (igual que rebuild)
    CloseExistingAccessProcesses
    
    ' PASO 2: Abrir Access para eliminar módulos específicos
    WScript.Echo "Paso 1: Eliminando modulos especificos..."
    
    Dim objAccess
    Set objAccess = OpenAccess(dbPath, GetDatabasePassword(dbPath))
    If objAccess Is Nothing Then
        WScript.Echo "❌ Error: No se pudo abrir Access"
        UpdateModules = False
        Exit Function
    End If
    
    ' Eliminar solo los módulos especificados
    Dim vbProject, vbComponent, componentCount, j
    Set vbProject = objAccess.VBE.ActiveVBProject
    componentCount = vbProject.VBComponents.Count
    
    For i = 0 To UBound(list)
        name = Trim(list(i))
        If name <> "" Then
            ' Buscar y eliminar el módulo específico
            For j = componentCount To 1 Step -1
                Set vbComponent = vbProject.VBComponents(j)
                If vbComponent.Name = name And (vbComponent.Type = 1 Or vbComponent.Type = 2) Then
                    WScript.Echo "  Eliminando: " & vbComponent.Name
                    vbProject.VBComponents.Remove vbComponent
                    If Err.Number <> 0 Then
                        WScript.Echo "  ❌ Error eliminando " & vbComponent.Name & ": " & Err.Description
                        Err.Clear
                    Else
                        WScript.Echo "  ✓ Eliminado: " & vbComponent.Name
                    End If
                    Exit For
                End If
            Next
        End If
    Next
    
    ' PASO 3: Cerrar y guardar (igual que rebuild)
    WScript.Echo "Paso 2: Cerrando base de datos..."
    objAccess.Quit 1  ' acQuitSaveAll = 1
    
    If Err.Number <> 0 Then
        WScript.Echo "Advertencia al cerrar Access: " & Err.Description
        Err.Clear
    End If
    
    Set objAccess = Nothing
    WScript.Echo "✓ Base de datos cerrada y guardada"
    
    ' PASO 4: Reabrir Access (igual que rebuild)
    WScript.Echo "Paso 3: Reabriendo base de datos..."
    
    Set objAccess = CreateObject("Access.Application")
    
    If Err.Number <> 0 Then
        WScript.Echo "❌ Error al crear nueva instancia de Access: " & Err.Description
        UpdateModules = False
        Exit Function
    End If
    
    ' Configurar Access en modo silencioso (igual que rebuild)
     objAccess.Visible = False
     objAccess.UserControl = False
     
     On Error Resume Next
     objAccess.DoCmd.SetWarnings False
     objAccess.Application.Echo False
     objAccess.DisplayAlerts = False
     objAccess.Application.AutomationSecurity = 1
     objAccess.VBE.MainWindow.Visible = False
     
     ' Configuraciones SetOption críticas para evitar confirmaciones (igual que rebuild)
     objAccess.SetOption "Confirm Action Queries", False
     objAccess.SetOption "Confirm Document Deletions", False
     objAccess.SetOption "Confirm Record Changes", False
     objAccess.SetOption "Confirm Document Save", False
     objAccess.SetOption "Auto Syntax Check", False
     objAccess.SetOption "Show Status Bar", False
     objAccess.SetOption "Show Animations", False
     objAccess.SetOption "Confirm Design Changes", False
     objAccess.SetOption "Confirm Object Deletions", False
     objAccess.SetOption "Track Name AutoCorrect Info", False
     objAccess.SetOption "Perform Name AutoCorrect", False
     objAccess.SetOption "Auto Compact", False
     
     Err.Clear
     On Error GoTo 0
    
    ' Abrir base de datos
    Dim strDbPassword
    strDbPassword = GetDatabasePassword(dbPath)
    
    If strDbPassword = "" Then
        objAccess.OpenCurrentDatabase dbPath
    Else
        objAccess.OpenCurrentDatabase dbPath, , strDbPassword
    End If
    
    If Err.Number <> 0 Then
        LogMessage "Error al reabrir base de datos: " & Err.Description
        UpdateModules = False
        Exit Function
    End If
    
    WScript.Echo "✓ Base de datos reabierta"
    
    ' PASO 5: Importar módulos específicos (igual que rebuild)
    WScript.Echo "Paso 4: Importando modulos especificos..."
    
    For i = 0 To UBound(list)
        name = Trim(list(i))
        If name <> "" Then
            LogMessage "Procesando modulo: " & name
            
            ' Usar la misma función que rebuild
            Dim importResult
            importResult = RebuildLike_ImportOne(objAccess, name)
            
            If Not importResult Then
                LogMessage "Error al importar modulo " & name
            End If
        End If
    Next
    
    ' PASO 6: Guardar módulos individualmente (igual que rebuild)
    LogMessage "Guardando modulos individualmente..."
    On Error Resume Next
    
    For Each vbComponent In objAccess.VBE.ActiveVBProject.VBComponents
        ' Solo guardar los módulos que acabamos de importar
        For i = 0 To UBound(list)
            name = Trim(list(i))
            If vbComponent.Name = name Then
                If vbComponent.Type = 1 Then  ' vbext_ct_StdModule
                    LogMessage "Guardando modulo: " & vbComponent.Name
                    objAccess.DoCmd.Save 5, vbComponent.Name  ' acModule = 5
                    If Err.Number <> 0 Then
                        LogMessage "Advertencia al guardar " & vbComponent.Name & ": " & Err.Description
                        Err.Clear
                    End If
                ElseIf vbComponent.Type = 2 Then  ' vbext_ct_ClassModule
                    LogMessage "Guardando clase: " & vbComponent.Name
                    objAccess.DoCmd.Save 7, vbComponent.Name  ' acClassModule = 7
                    If Err.Number <> 0 Then
                        LogMessage "Advertencia al guardar " & vbComponent.Name & ": " & Err.Description
                        Err.Clear
                    End If
                End If
                Exit For
            End If
        Next
    Next
    
    WScript.Echo "=== ACTUALIZACION COMPLETADA EXITOSAMENTE ==="
    LogMessage "Los modulos especificados han sido actualizados"
    
    ' Cerrar Access
    objAccess.Quit
    Set objAccess = Nothing
    
    UpdateModules = True
End Function

' Función para listar objetos de la base de datos
Sub ListObjects(dbPath, showSchema, outputToFile)
    LogMessage "Listando objetos de: " & objFSO.GetFileName(dbPath)
    
    Set objAccess = OpenAccess(dbPath, gPassword)
    If objAccess Is Nothing Then
        Exit Sub
    End If
    
    ' Configurar salida
    Dim outputFile, outputPath
    If outputToFile Then
        Dim dbName
        dbName = objFSO.GetBaseName(dbPath)
        outputPath = objFSO.GetAbsolutePathName(".") & "\" & dbName & "_listobjects.txt"
        Set outputFile = objFSO.CreateTextFile(outputPath, True)
        WScript.Echo "Exportando resultados a: " & outputPath
    End If

    WScript.Echo "=== TABLAS ==="
    If outputToFile Then outputFile.WriteLine "=== TABLAS ==="
    
    Dim tbl, tableCount
    tableCount = 0
    For Each tbl In objAccess.CurrentDb.TableDefs
        If Left(tbl.Name, 4) <> "MSys" And Left(tbl.Name, 1) <> "~" Then
            tableCount = tableCount + 1
            If showSchema Then
                WScript.Echo "  " & tbl.Name & " (" & tbl.RecordCount & " registros)"
                If outputToFile Then outputFile.WriteLine "  " & tbl.Name & " (" & tbl.RecordCount & " registros)"
                
                ' Mostrar campos si se solicita esquema
                Dim fld
                For Each fld In tbl.Fields
                    WScript.Echo "    - " & fld.Name & " (" & GetFieldTypeName(fld.Type) & ")"
                    If outputToFile Then outputFile.WriteLine "    - " & fld.Name & " (" & GetFieldTypeName(fld.Type) & ")"
                Next
            Else
                WScript.Echo "  " & tbl.Name
                If outputToFile Then outputFile.WriteLine "  " & tbl.Name
            End If
        End If
    Next

    WScript.Echo "=== FORMULARIOS ==="
    If outputToFile Then outputFile.WriteLine "=== FORMULARIOS ==="
    
    Dim frm, formCount
    formCount = 0
    For Each frm In objAccess.CurrentProject.AllForms
        formCount = formCount + 1
        WScript.Echo "  " & frm.Name
        If outputToFile Then outputFile.WriteLine "  " & frm.Name
    Next

    WScript.Echo "=== CONSULTAS ==="
    If outputToFile Then outputFile.WriteLine "=== CONSULTAS ==="
    
    Dim qry, queryCount
    queryCount = 0
    For Each qry In objAccess.CurrentDb.QueryDefs
        queryCount = queryCount + 1
        WScript.Echo "  " & qry.Name
        If outputToFile Then outputFile.WriteLine "  " & qry.Name
    Next

    WScript.Echo "=== MODULOS ==="
    If outputToFile Then outputFile.WriteLine "=== MODULOS ==="
    
    Dim mdl, moduleCount
    moduleCount = 0
    For Each mdl In objAccess.CurrentProject.AllModules
        moduleCount = moduleCount + 1
        WScript.Echo "  " & mdl.Name
        If outputToFile Then outputFile.WriteLine "  " & mdl.Name
    Next
    
    WScript.Echo ""
    WScript.Echo "RESUMEN:"
    WScript.Echo "  Tablas: " & tableCount
    WScript.Echo "  Formularios: " & formCount
    WScript.Echo "  Consultas: " & queryCount
    WScript.Echo "  Modulos: " & moduleCount
    
    If outputToFile Then
        outputFile.WriteLine ""
        outputFile.WriteLine "RESUMEN:"
        outputFile.WriteLine "  Tablas: " & tableCount
        outputFile.WriteLine "  Formularios: " & formCount
        outputFile.WriteLine "  Consultas: " & queryCount
        outputFile.WriteLine "  Modulos: " & moduleCount
        outputFile.Close
        Set outputFile = Nothing
        WScript.Echo "Archivo generado exitosamente: " & outputPath
    End If

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
    objAccess.DoCmd.SetWarnings False
    objAccess.Application.Echo False
    objAccess.DisplayAlerts = False
    ' Configuraciones adicionales para suprimir diálogos
    objAccess.Application.AutomationSecurity = 1  ' msoAutomationSecurityLow
    objAccess.VBE.MainWindow.Visible = False
    
    ' Configuraciones SetOption críticas para evitar confirmaciones
    objAccess.SetOption "Confirm Action Queries", False
    objAccess.SetOption "Confirm Document Deletions", False
    objAccess.SetOption "Confirm Record Changes", False
    objAccess.SetOption "Confirm Document Save", False
    objAccess.SetOption "Auto Syntax Check", False
    objAccess.SetOption "Show Status Bar", False
    objAccess.SetOption "Show Animations", False
    objAccess.SetOption "Confirm Design Changes", False
    objAccess.SetOption "Confirm Object Deletions", False
    objAccess.SetOption "Track Name AutoCorrect Info", False
    objAccess.SetOption "Perform Name AutoCorrect", False
    objAccess.SetOption "Auto Compact", False
    
    Err.Clear
    On Error GoTo 0
    
    ' Determinar contraseña para la base de datos
    Dim strDbPassword
    strDbPassword = GetDatabasePassword(gDbPath)
    
    ' Abrir base de datos
    If strDbPassword = "" Then
        objAccess.OpenCurrentDatabase gDbPath
    Else
        objAccess.OpenCurrentDatabase gDbPath, , strDbPassword
    End If
    
    If Err.Number <> 0 Then
        LogMessage "Error al reabrir base de datos: " & Err.Description
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
    
    ' PASO 4.3: Guardar cada modulo individualmente
    LogMessage "Guardando modulos individualmente..."
    On Error Resume Next
    
    For Each vbComponent In objAccess.VBE.ActiveVBProject.VBComponents
        If vbComponent.Type = 1 Then  ' vbext_ct_StdModule
            LogMessage "Guardando modulo: " & vbComponent.Name
            objAccess.DoCmd.Save 5, vbComponent.Name  ' acModule = 5
            If Err.Number <> 0 Then
                LogMessage "Advertencia al guardar " & vbComponent.Name & ": " & Err.Description
                Err.Clear
            End If
        ElseIf vbComponent.Type = 2 Then  ' vbext_ct_ClassModule
            LogMessage "Guardando clase: " & vbComponent.Name
            objAccess.DoCmd.Save 7, vbComponent.Name  ' acClassModule = 7
            If Err.Number <> 0 Then
                LogMessage "Advertencia al guardar " & vbComponent.Name & ": " & Err.Description
                Err.Clear
            End If
        End If
    Next
    
    ' PASO 4.4: Verificacion de integridad y compilacion
    LogMessage "Verificando integridad de nombres de modulos..."
    Call VerifyModuleNames(objAccess, g_ModulesSrcPath)
    
    LogMessage "=== RECONSTRUCCION COMPLETADA EXITOSAMENTE ==="
    LogMessage "El proyecto VBA ha sido completamente reconstruido"
    LogMessage "Todos los modulos han sido reimportados desde /src"
    
    On Error GoTo 0
End Sub

' ============================================================================
' SECCIÓN 8: FUNCIONES UI AS CODE - PRINCIPALES
' ============================================================================

Function UI_RebuildAll(dbPath, files, config)
    On Error Resume Next
    UI_RebuildAll = False
    
    ' Validar parámetros de entrada
    If Len(dbPath) = 0 Then
        LogMessage "UI_RebuildAll: dbPath no puede estar vacío"
        Exit Function
    End If
    
    If Not IsArray(files) Then
        LogMessage "UI_RebuildAll: files debe ser un array"
        Exit Function
    End If
    
    ' Intentar la función completa con manejo de errores mejorado
    Dim result
    result = UI_RebuildAll_Complete(dbPath, files, config)
    
    If Err.Number <> 0 Then
        LogMessage "UI_RebuildAll: Error " & Err.Number & " - " & Err.Description
        Err.Clear
        UI_RebuildAll = False
    Else
        UI_RebuildAll = result
    End If
    
    On Error GoTo 0
End Function

Function UI_RebuildAll_Complete(dbPath, files, config)
    On Error Resume Next
    UI_RebuildAll_Complete = False
    
    ' Abrir Access
    Dim app
    Set app = OpenAccess(dbPath, gPassword)
    If app Is Nothing Then
        LogMessage "UI_RebuildAll_Complete: No se pudo abrir Access"
        Exit Function
    End If
    
    ' Validacion fail-fast para VBE antes de continuar
    If Not CheckVBProjectAccess(app) Then
        LogError "ui-rebuild: VBE no disponible - operacion abortada"
        CloseAccess app
        UI_RebuildAll_Complete = False
        Exit Function
    End If
    
    ' Aplicar configuracion anti-UI
    AntiUI app
    
    ' Procesar cada archivo
    Dim i, success, allSuccess
    allSuccess = True
    
    For i = 0 To UBound(files)
        Dim fileObj: Set fileObj = CreateObject("Scripting.Dictionary")
        fileObj.Add "Path", files(i)
        fileObj.Add "Name", objFSO.GetFileName(files(i))
        
        success = ProcessSingleFormFile(app, fileObj, objFSO)
        If Not success Then
            LogMessage "UI_RebuildAll_Complete: Falló procesamiento de " & fileObj("Name")
            allSuccess = False
        End If
    Next
    
    ' Verificar responsividad tras operaciones
    If Not EnsureResponsive(app, "UI_RebuildAll_Complete final") Then
        LogMessage "ADVERTENCIA: Access no responde correctamente al final de UI_RebuildAll_Complete"
    End If
    
    ' Cerrar Access
    CloseAccess app
    
    UI_RebuildAll_Complete = allSuccess
    On Error GoTo 0
End Function

Function UI_UpdateSome(dbPath, items, config)
    On Error Resume Next
    UI_UpdateSome = False
    
    ' Validar parámetros de entrada
    If Len(dbPath) = 0 Then
        LogMessage "UI_UpdateSome: dbPath no puede estar vacío"
        Exit Function
    End If
    
    If Not IsArray(items) Then
        LogMessage "UI_UpdateSome: items debe ser un array"
        Exit Function
    End If
    
    ' Abrir Access
    Dim app
    Set app = OpenAccess(dbPath, gPassword)
    If app Is Nothing Then
        LogMessage "UI_UpdateSome: No se pudo abrir Access"
        Exit Function
    End If
    
    ' Validacion fail-fast para VBE antes de continuar
    If Not CheckVBProjectAccess(app) Then
        LogError "ui-update: VBE no disponible - operacion abortada"
        CloseAccess app
        UI_UpdateSome = False
        Exit Function
    End If
    
    ' Procesar cada item usando UI_UpdateSome_ByFile
    Dim i, success, allSuccess
    allSuccess = True
    
    For i = 0 To UBound(items)
        Dim item: item = items(i)
        Dim jsonPath, targetName
        
        ' Extraer jsonPath y targetName del item
        If IsObject(item) Then
            jsonPath = item.jsonPath
            targetName = item.targetName
        ElseIf InStr(item, "|") > 0 Then
            ' Formato "jsonPath|targetName"
            Dim parts: parts = Split(item, "|")
            jsonPath = parts(0)
            targetName = parts(1)
        Else
            ' Solo jsonPath, extraer targetName del nombre del archivo
            jsonPath = item
            targetName = Replace(Replace(objFSO.GetBaseName(jsonPath), "_", ""), " ", "")
        End If
        
        success = UI_UpdateSome_ByFile(app, jsonPath, targetName)
        If Not success Then
            LogMessage "UI_UpdateSome: Falló procesamiento de " & jsonPath
            allSuccess = False
        End If
    Next
    
    ' Cerrar Access
    CloseAccess app
    
    UI_UpdateSome = allSuccess
    On Error GoTo 0
End Function

Function UI_UpdateSome_ByFile(app, jsonPath, targetName)
    On Error Resume Next
    UI_UpdateSome_ByFile = False
    
    ' Leer el archivo JSON
    Dim jsonContent: jsonContent = ReadAllText(jsonPath)
    If Len(jsonContent) = 0 Then
        LogMessage "ui-update-some: no se pudo leer " & jsonPath
        Exit Function
    End If
    
    ' Parsear JSON
    Dim root: Set root = JsonParse(jsonContent)
    If root Is Nothing Then
        LogMessage "ui-update-some: JSON invalido en " & jsonPath
        Exit Function
    End If
    
    ' Aplicar configuracion anti-UI
    AntiUI app
    
    ' Crear formulario temporal
    app.DoCmd.CreateForm
    If FailIfErr("CreateForm en UI_UpdateSome_ByFile") Then Exit Function
    
    ' Verificar responsividad tras CreateForm
    If Not EnsureResponsive(app, "UI_UpdateSome_ByFile post-CreateForm") Then
        LogMessage "ADVERTENCIA: Access no responde correctamente tras CreateForm en UI_UpdateSome_ByFile"
    End If
    
    ' Aplicar configuracion anti-UI despues de crear
    AntiUI app
    
    ' Obtener nombre del formulario creado con reintento
    Dim tmpName
    If Not WaitActiveFormName(app, tmpName) Then
        LogMessage "ui-update-some: no se pudo obtener ActiveForm tras CreateForm (timeout)"
        Exit Function
    End If
    
    ' Renombrar a temporal
    Dim tempFormName: tempFormName = targetName & "__tmp"
    
    ' Aplicar configuracion anti-UI antes de renombrar
    AntiUI app
    
    If Not DoCmdSafe_Rename(app, tempFormName, acForm, tmpName, "UI_UpdateSome_ByFile rename to temp") Then
        LogMessage "ui-update-some: no se pudo renombrar a temporal"
        Exit Function
    End If
    
    ' Verificar que el renombrado fue exitoso
    Dim activeName
    If Not WaitActiveFormName(app, activeName) Or LCase(activeName) <> LCase(tempFormName) Then
        LogError "ui-update-some: formulario no se renombró correctamente a " & tempFormName
        Exit Function
    End If
    
    ' Aplicar configuracion anti-UI despues de renombrar
    AntiUI app
    
    ' Cerrar el formulario temporal SIN guardar cambios (ya compilado globalmente)
    app.DoCmd.Close acForm, tempFormName, acSaveNo ' acForm, acSaveNo
    
    ' Verificar responsividad tras Close
    If Not EnsureResponsive(app, "UI_UpdateSome_ByFile post-Close") Then
        LogMessage "ADVERTENCIA: Access no responde correctamente tras Close en UI_UpdateSome_ByFile"
    End If
    
    ' Swap atomico usando SafeSwapForm
    SafeSwapForm app, targetName, tempFormName
    
    ' Verificar responsividad tras operacion critica
    If Not EnsureResponsive(app, "UI_UpdateSome_ByFile post-swap") Then
        LogMessage "ADVERTENCIA: Access no responde correctamente tras SafeSwapForm en UI_UpdateSome_ByFile"
    End If
    
    UI_UpdateSome_ByFile = True
End Function

Function UI_ImportOne(app, jsonPath, formName)
    On Error Resume Next
    UI_ImportOne = False
    
    ' Leer el archivo JSON
    Dim jsonContent: jsonContent = ReadAllText(jsonPath)
    If Len(jsonContent) = 0 Then
        LogMessage "ui-import: no se pudo leer " & jsonPath
        Exit Function
    End If
    
    ' Parsear JSON
    Dim root: Set root = JsonParse(jsonContent)
    If root Is Nothing Then
        LogMessage "ui-import: JSON invalido en " & jsonPath
        Exit Function
    End If
    
    ' Validar estructura JSON
    Dim validationError: validationError = ""
    If Not ValidateFormJson(root, validationError) Then
        LogMessage "ui-import: JSON invalido - " & validationError
        Exit Function
    End If

    ' Patron atomico: crear formulario temporal
    Dim tmpName: tmpName = formName & "__tmp"
    
    ' Cerrar cualquier formulario activo en diseño de ejecuciones previas (silencioso)
    On Error Resume Next
    AntiUI app
    app.DoCmd.Close acForm, app.Screen.ActiveForm.Name, acSaveNo
    Err.Clear
    
    ' Crear temporal
    On Error Resume Next
    AntiUI app
    app.DoCmd.CreateForm
    If Err.Number <> 0 Then 
        LogMessage "ui-import: no se pudo crear temporal: " & Err.Description
        Exit Function
    End If
    
    ' Verificar responsividad tras CreateForm
    If Not EnsureResponsive(app, "UI_ImportOne post-CreateForm") Then
        LogMessage "ADVERTENCIA: Access no responde correctamente tras CreateForm en UI_ImportOne"
    End If
    
    Err.Clear
    
    ' Aplicar configuracion anti-UI despues de crear
    AntiUI app
    
    ' Obtener nombre del formulario creado con timeout configurable
    Dim created
    If Not WaitActiveFormName(app, created) Then
        LogMessage "ui-import: no se pudo obtener ActiveForm tras CreateForm"
        Exit Function
    End If
    On Error GoTo 0
    
    ' Reforzar blindaje antes del Rename
    AntiUI app
    
    ' Renombrar temporal con anti-UI
    If Not DoCmdSafe_Rename(app, tmpName, acForm, created, "UI_ImportOne rename temp") Then
        LogMessage "ui-import: no se pudo renombrar temporal"
        Exit Function
    End If
    
    ' Verificar que el renombrado fue exitoso
    Dim activeName4
    If Not WaitActiveFormName(app, activeName4) Or LCase(activeName4) <> LCase(tmpName) Then
        LogError "ui-import: formulario no se renombró correctamente a " & tmpName
        Exit Function
    End If
    
    ' Abrir en modo diseño con anti-UI
    If Not DoCmdSafe_OpenFormDesignHidden(app, tmpName, "UI_ImportOne open temp") Then
        LogMessage "ui-import: no se pudo abrir temporal en diseño"
        Exit Function
    End If
    
    Dim frm: Set frm = app.Forms(tmpName)

    ' Aplicar propiedades y controles sobre el temporal
    SetPropIfExists frm, "RecordSource", GetDict(root,"recordSource","")
    SetPropIfExists frm, "Caption",      GetDict(root,"caption","")
    ApplySections frm, root

    ' Construir diccionario de handlers desde JSON
    Dim handlers: Set handlers = CreateObject("Scripting.Dictionary")
    If root.Exists("code") Then
        Dim codeSection: Set codeSection = root("code")
        If codeSection.Exists("module") Then
            Dim moduleSection: Set moduleSection = codeSection("module")
            If moduleSection.Exists("handlers") And IsArray(moduleSection("handlers")) Then
                Dim handlersArray, h, handlerKey
                handlersArray = moduleSection("handlers")
                For h = 0 To UBound(handlersArray)
                    If IsObject(handlersArray(h)) Then
                        Dim handlerObj: Set handlerObj = handlersArray(h)
                        If handlerObj.Exists("control") And handlerObj.Exists("event") Then
                            handlerKey = handlerObj("control") & "." & handlerObj("event")
                            handlers.Add handlerKey, True
                        End If
                    End If
                Next
            End If
        End If
    End If

    ' Controles
    DeleteAllControls app, tmpName, frm
    Dim arr, i, cnt
    If root.Exists("controls") Then
        arr = root("controls")
        If IsArray(arr) Then
            For i = 0 To UBound(arr)
                If IsObject(arr(i)) Then
                    cnt = CreateControlFromJson(app, frm, arr(i), handlers)
                    If cnt Is Nothing Then
                        LogMessage "ui-import: fallo al crear control " & i
                    End If
                End If
            Next
        End If
    End If
    
    ' Cerrar formulario temporal
    app.DoCmd.Close acForm, tmpName, acSaveYes
    
    ' Swap atomico
    SafeSwapForm app, formName, tmpName
    
    UI_ImportOne = True
    On Error GoTo 0
End Function

Function ProcessSingleFormFile(app, file, fso)
    On Error Resume Next
    ProcessSingleFormFile = False
    
    Dim baseName: baseName = fso.GetBaseName(file("Name"))
    Dim jsonPath: jsonPath = file("Path")
    
    LogMessage "ui-rebuild-all: procesando " & baseName & " desde " & jsonPath
    
    ' Leer el archivo JSON
    Dim jsonContent: jsonContent = ReadAllText(jsonPath)
    If Len(jsonContent) = 0 Then
        LogMessage "ui-rebuild-all: no se pudo leer " & jsonPath
        Exit Function
    End If
    
    ' Parsear JSON
    Dim root: Set root = JsonParse(jsonContent)
    If root Is Nothing Then
        LogMessage "ui-rebuild-all: JSON invalido en " & jsonPath
        Exit Function
    End If
    
    ' Aplicar configuracion anti-UI
    AntiUI app
    
    ' Crear formulario temporal
    app.DoCmd.CreateForm
    If FailIfErr("CreateForm en ProcessSingleFormFile") Then Exit Function
    
    ' Verificar responsividad tras CreateForm
    If Not EnsureResponsive(app, "ProcessSingleFormFile post-CreateForm") Then
        LogMessage "ADVERTENCIA: Access no responde correctamente tras CreateForm en ProcessSingleFormFile"
    End If
    
    ' Aplicar configuracion anti-UI despues de crear
    AntiUI app
    
    ' Obtener nombre del formulario creado con timeout configurable
    Dim tmpName
    If Not WaitActiveFormName(app, tmpName) Then
        LogMessage "ui-rebuild-all: no se pudo obtener ActiveForm tras CreateForm para " & baseName
        Exit Function
    End If
    
    ' Renombrar a temporal
    Dim tempFormName: tempFormName = baseName & "__tmp"
    
    If Not DoCmdSafe_Rename(app, tempFormName, acForm, tmpName, "ProcessSingleFormFile rename") Then
        LogMessage "ui-rebuild-all: no se pudo renombrar a temporal"
        Exit Function
    End If
    
    ' Aplicar propiedades del formulario
    Dim frm: Set frm = app.Forms(tempFormName)
    SetPropIfExists frm, "RecordSource", GetDict(root,"recordSource","")
    SetPropIfExists frm, "Caption", GetDict(root,"caption","")
    
    ' Aplicar controles
    DeleteAllControls app, tempFormName, frm
    If root.Exists("controls") Then
        Dim arr: arr = root("controls")
        If IsArray(arr) Then
            Dim i, cnt
            For i = 0 To UBound(arr)
                If IsObject(arr(i)) Then
                    cnt = CreateControlFromJson(app, frm, arr(i), Nothing)
                End If
            Next
        End If
    End If
    
    ' Cerrar formulario temporal
    app.DoCmd.Close acForm, tempFormName, acSaveYes
    
    ' Swap atomico
    SafeSwapForm app, baseName, tempFormName
    
    ProcessSingleFormFile = True
    On Error GoTo 0
End Function

' ============================================================================
' SECCIÓN 8.1: FUNCIONES DE DETECCIÓN DE MÓDULOS UI AS CODE
' ============================================================================

' FUNCION: UI_DetectModuleAndHandlers
' Descripcion: Detecta modulo de formulario y sus handlers de eventos
' Parametros: jsonWriter - objeto para escribir JSON
'            formName - nombre del formulario
'            srcDir - directorio fuente donde buscar modulos
'            verbose - mostrar informacion detallada
' Retorna: Scripting.Dictionary con claves "Control.Event"
' ===================================================================
Function UI_DetectModuleAndHandlers(jsonWriter, formName, srcDir, verbose)
    Dim objFSO, moduleFile, moduleExists, moduleFilename
    Dim handlers, fileContent, regEx, matches, match
    Dim i, controlName, eventName, signature
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set handlers = CreateObject("Scripting.Dictionary")
    
    moduleExists = False
    moduleFilename = ""
    
    ' Heuristica para encontrar el archivo del modulo
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
    
    ' Escribir seccion code.module
    jsonWriter.WriteProperty "code", ""
    jsonWriter.StartObject
    jsonWriter.WriteProperty "module", ""
    jsonWriter.StartObject
    jsonWriter.WriteProperty "exists", moduleExists
    jsonWriter.WriteProperty "filename", moduleFilename
    
    ' Si existe el modulo, parsear handlers
    If moduleExists Then
        If verbose Then LogMessage "Detectando handlers en: " & moduleFile
        
        ' Leer contenido del archivo
        Dim textStream
        Set textStream = objFSO.OpenTextFile(moduleFile, 1) ' ForReading
        fileContent = textStream.ReadAll
        textStream.Close
        
        ' Crear expresion regular para detectar handlers
        Set regEx = CreateObject("VBScript.RegExp")
        regEx.Global = True
        regEx.IgnoreCase = True
        regEx.MultiLine = True
        regEx.Pattern = "^\s*(Public|Private)?\s*Sub\s+([A-Za-z0-9_]+)_(Click|DblClick|Current|Load|Open|GotFocus|LostFocus|Change|AfterUpdate|BeforeUpdate)\s*\("
        
        Set matches = regEx.Execute(fileContent)
        
        ' Procesar matches
        jsonWriter.WriteProperty "handlers", ""
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
            
            If verbose Then LogMessage "Handler detectado: " & controlName & "." & eventName
        Next
        
        jsonWriter.EndArray ' handlers
    Else
        ' No hay modulo, array vacio
        jsonWriter.WriteProperty "handlers", ""
        jsonWriter.StartArray
        jsonWriter.EndArray
        
        If verbose Then LogMessage "No se encontro modulo para el formulario " & formName
    End If
    
    jsonWriter.EndObject ' module
    jsonWriter.EndObject ' code
    
    ' Retornar diccionario de handlers para uso posterior
    Set UI_DetectModuleAndHandlers = handlers
End Function

Function UI_DetectModuleAndHandlers_Safe(formName, jsonWriter)
    On Error Resume Next
    UI_DetectModuleAndHandlers_Safe = False
    
    If Len(formName) = 0 Or Not IsObject(jsonWriter) Then
        LogMessage "UI_DetectModuleAndHandlers_Safe: parametros invalidos"
        Exit Function
    End If
    
    ' Intentar la función original con manejo de errores
    Dim result
    result = UI_DetectModuleAndHandlers(formName, jsonWriter)
    
    If Err.Number <> 0 Then
        LogMessage "UI_DetectModuleAndHandlers_Safe: Error " & Err.Number & " - " & Err.Description
        Err.Clear
        UI_DetectModuleAndHandlers_Safe = False
    Else
        UI_DetectModuleAndHandlers_Safe = result
    End If
End Function

' ===== FUNCIONES PARA EXPORT-FORM =====

Sub ExportFormToJson(dbPath, formName, outputPath, password)
    Dim objAccess, frm, jsonContent
    
    ' SONDA 1: Verificar parámetros de entrada
    If gVerbose Then WScript.Echo "[SONDA] Iniciando ExportFormToJson con parametros:"
    If gVerbose Then WScript.Echo "[SONDA]   - dbPath: " & dbPath
    If gVerbose Then WScript.Echo "[SONDA]   - formName: " & formName
    If gVerbose Then WScript.Echo "[SONDA]   - outputPath: " & outputPath
    If gVerbose Then WScript.Echo "[SONDA]   - password: " & IIf(password = "", "(vacio)", "(proporcionada)")
    
    ' SONDA 2: Verificar existencia del archivo de base de datos
    If gVerbose Then WScript.Echo "[SONDA] Verificando existencia del archivo de base de datos..."
    If Not objFSO.FileExists(dbPath) Then
        WScript.Echo "[ERROR] El archivo de base de datos no existe: " & dbPath
        WScript.Quit 1
    End If
    If gVerbose Then WScript.Echo "[SONDA] Archivo de base de datos encontrado correctamente"
    
    ' SONDA 3: Intentar abrir Access
    If gVerbose Then WScript.Echo "[SONDA] Intentando abrir Access..."
    Set objAccess = OpenAccess(dbPath, password)
    If objAccess Is Nothing Then
        WScript.Echo "[ERROR] No se pudo abrir la base de datos"
        WScript.Echo "[SONDA] Posibles causas:"
        WScript.Echo "[SONDA]   - Contraseña incorrecta"
        WScript.Echo "[SONDA]   - Archivo corrupto"
        WScript.Echo "[SONDA]   - Permisos insuficientes"
        WScript.Quit 1
    End If
    If gVerbose Then WScript.Echo "[SONDA] Access abierto correctamente"
    
    On Error Resume Next
    
    ' SONDA 4: Construir ruta de salida por defecto si no se especifica
    If outputPath = "" Then
        If gVerbose Then WScript.Echo "[SONDA] Construyendo ruta de salida por defecto..."
        Dim uiFormsPath
        uiFormsPath = objFSO.GetParentFolderName(WScript.ScriptFullName) & "\ui\forms\"
        If gVerbose Then WScript.Echo "[SONDA] Ruta calculada: " & uiFormsPath
        
        ' Asegurarse de que el directorio existe
        If Not objFSO.FolderExists(uiFormsPath) Then 
            If gVerbose Then WScript.Echo "[SONDA] Creando directorio ui\forms..."
            If Not objFSO.FolderExists(objFSO.GetParentFolderName(WScript.ScriptFullName) & "\ui") Then
                objFSO.CreateFolder objFSO.GetParentFolderName(WScript.ScriptFullName) & "\ui"
                If gVerbose Then WScript.Echo "[SONDA] Directorio ui creado"
            End If
            objFSO.CreateFolder uiFormsPath
            If gVerbose Then WScript.Echo "[SONDA] Directorio forms creado"
        End If
        outputPath = objFSO.BuildPath(uiFormsPath, formName & ".json")
        If gVerbose Then WScript.Echo "[SONDA] Ruta final de salida: " & outputPath
    End If
    
    ' SONDA 5: Verificar que el formulario existe en la base de datos
    If gVerbose Then WScript.Echo "[SONDA] Verificando existencia del formulario en la base de datos..."
    Dim formExists: formExists = False
    Dim formObj
    For Each formObj In objAccess.CurrentProject.AllForms
        If LCase(formObj.Name) = LCase(formName) Then
            formExists = True
            Exit For
        End If
    Next
    
    If Not formExists Then
        WScript.Echo "[ERROR] El formulario '" & formName & "' no existe en la base de datos"
        WScript.Echo "[SONDA] Formularios disponibles:"
        For Each formObj In objAccess.CurrentProject.AllForms
            WScript.Echo "[SONDA]   - " & formObj.Name
        Next
        Call CloseAccess(objAccess)
        WScript.Quit 1
    End If
    If gVerbose Then WScript.Echo "[SONDA] Formulario encontrado en la base de datos"
    
    ' SONDA 6: Abrir formulario en modo diseño y oculto
    If gVerbose Then WScript.Echo "[SONDA] Abriendo formulario '" & formName & "' en modo diseño..."
    objAccess.DoCmd.OpenForm formName, acViewDesign, , , , acHidden
    
    If Err.Number <> 0 Then
        WScript.Echo "[ERROR] No se pudo abrir el formulario '" & formName & "': " & Err.Description
        WScript.Echo "[SONDA] Codigo de error: " & Err.Number
        WScript.Echo "[SONDA] Posibles causas:"
        WScript.Echo "[SONDA]   - Formulario corrupto"
        WScript.Echo "[SONDA]   - Dependencias faltantes"
        WScript.Echo "[SONDA]   - Permisos insuficientes"
        Call CloseAccess(objAccess)
        WScript.Quit 1
    End If
    If gVerbose Then WScript.Echo "[SONDA] Formulario abierto correctamente en modo diseño"
    
    ' SONDA 7: Obtener referencia al formulario
    If gVerbose Then WScript.Echo "[SONDA] Obteniendo referencia al formulario..."
    Set frm = objAccess.Forms(formName)
    If frm Is Nothing Then
        WScript.Echo "[ERROR] No se pudo obtener referencia al formulario"
        Call CloseAccess(objAccess)
        WScript.Quit 1
    End If
    If gVerbose Then WScript.Echo "[SONDA] Referencia al formulario obtenida correctamente"
    If gVerbose Then WScript.Echo "[SONDA] Nombre del formulario: " & frm.Name
    If gVerbose Then WScript.Echo "[SONDA] Numero de controles: " & frm.Controls.Count
    
    ' SONDA 8: Generar JSON básico del formulario
    If gVerbose Then WScript.Echo "[SONDA] Generando JSON del formulario..."
    jsonContent = GenerateFormJson(frm)
    If Len(jsonContent) = 0 Then
        WScript.Echo "[ERROR] No se pudo generar el contenido JSON"
        objAccess.DoCmd.Close acForm, formName, acSaveNo
        Call CloseAccess(objAccess)
        WScript.Quit 1
    End If
    If gVerbose Then WScript.Echo "[SONDA] JSON generado correctamente (" & Len(jsonContent) & " caracteres)"
    
    ' SONDA 9: Cerrar formulario sin guardar cambios
    If gVerbose Then WScript.Echo "[SONDA] Cerrando formulario..."
    objAccess.DoCmd.Close acForm, formName, acSaveNo
    If gVerbose Then WScript.Echo "[SONDA] Formulario cerrado correctamente"
    
    ' SONDA 10: Cerrar Access
    If gVerbose Then WScript.Echo "[SONDA] Cerrando Access..."
    Call CloseAccess(objAccess)
    If gVerbose Then WScript.Echo "[SONDA] Access cerrado correctamente"
    
    ' SONDA 11: Escribir archivo JSON
    If gVerbose Then WScript.Echo "[SONDA] Escribiendo archivo JSON a: " & outputPath
    Call WriteTextFile(outputPath, jsonContent)
    
    ' SONDA 12: Verificar que el archivo se escribió correctamente
    If objFSO.FileExists(outputPath) Then
        Dim fileSize: fileSize = objFSO.GetFile(outputPath).Size
        If gVerbose Then WScript.Echo "[SONDA] Archivo escrito correctamente (" & fileSize & " bytes)"
        WScript.Echo "Formulario exportado exitosamente a: " & outputPath
    Else
        WScript.Echo "[ERROR] No se pudo escribir el archivo JSON"
        WScript.Quit 1
    End If
    
    On Error GoTo 0
End Sub

Function GenerateFormJson(frm)
    Dim json, ctrl, i
    
    ' SONDA JSON 1: Iniciar generación de JSON
    If gVerbose Then WScript.Echo "[SONDA JSON] Iniciando generacion de JSON para formulario: " & frm.Name
    
    json = "{" & vbCrLf
    json = json & "  ""schemaVersion"": ""1.0.0""," & vbCrLf
    json = json & "  ""formName"": """ & frm.Name & """," & vbCrLf
    json = json & "  ""properties"": {" & vbCrLf
    
    On Error Resume Next
    
    ' SONDA JSON 2: Extraer propiedades básicas del formulario
    If gVerbose Then WScript.Echo "[SONDA JSON] Extrayendo propiedades basicas del formulario..."
    
    ' Propiedades básicas del formulario
    json = json & "    ""caption"": """ & EscapeJsonString(frm.Caption) & """," & vbCrLf
    If Err.Number <> 0 Then
        If gVerbose Then WScript.Echo "[SONDA JSON] Error obteniendo Caption: " & Err.Description
        Err.Clear
    End If
    
    json = json & "    ""width"": " & frm.Width & "," & vbCrLf
    If Err.Number <> 0 Then
        If gVerbose Then WScript.Echo "[SONDA JSON] Error obteniendo Width: " & Err.Description
        Err.Clear
    End If
    
    json = json & "    ""height"": " & frm.Section(0).Height & "," & vbCrLf
    If Err.Number <> 0 Then
        If gVerbose Then WScript.Echo "[SONDA JSON] Error obteniendo Height: " & Err.Description
        Err.Clear
    End If
    
    json = json & "    ""recordSource"": """ & EscapeJsonString(frm.RecordSource) & """," & vbCrLf
    If Err.Number <> 0 Then
        If gVerbose Then WScript.Echo "[SONDA JSON] Error obteniendo RecordSource: " & Err.Description
        Err.Clear
    End If
    
    json = json & "    ""modal"": " & LCase(CStr(frm.Modal)) & "," & vbCrLf
    If Err.Number <> 0 Then
        If gVerbose Then WScript.Echo "[SONDA JSON] Error obteniendo Modal: " & Err.Description
        Err.Clear
    End If
    
    json = json & "    ""popUp"": " & LCase(CStr(frm.PopUp)) & vbCrLf
    If Err.Number <> 0 Then
        If gVerbose Then WScript.Echo "[SONDA JSON] Error obteniendo PopUp: " & Err.Description
        Err.Clear
    End If
    
    json = json & "  }," & vbCrLf
    json = json & "  ""controls"": [" & vbCrLf
    
    ' SONDA JSON 3: Procesar controles
    If gVerbose Then WScript.Echo "[SONDA JSON] Procesando " & frm.Controls.Count & " controles..."
    
    ' Procesar controles
    For i = 0 To frm.Controls.Count - 1
        If gVerbose Then WScript.Echo "[SONDA JSON] Procesando control " & (i + 1) & " de " & frm.Controls.Count
        
        Set ctrl = frm.Controls(i)
        If ctrl Is Nothing Then
            If gVerbose Then WScript.Echo "[SONDA JSON] Control " & i & " es Nothing, saltando..."
        Else
            If gVerbose Then WScript.Echo "[SONDA JSON] Control: " & ctrl.Name & " (Tipo: " & ctrl.ControlType & ")"
            
            If i > 0 Then json = json & "," & vbCrLf
            
            json = json & "    {" & vbCrLf
            json = json & "      ""name"": """ & EscapeJsonString(ctrl.Name) & """," & vbCrLf
            If Err.Number <> 0 Then
                If gVerbose Then WScript.Echo "[SONDA JSON] Error obteniendo Name del control " & i & ": " & Err.Description
                Err.Clear
            End If
            
            json = json & "      ""type"": """ & GetControlTypeName(ctrl.ControlType) & """," & vbCrLf
            If Err.Number <> 0 Then
                If gVerbose Then WScript.Echo "[SONDA JSON] Error obteniendo ControlType del control " & i & ": " & Err.Description
                Err.Clear
            End If
            
            json = json & "      ""left"": " & ctrl.Left & "," & vbCrLf
            If Err.Number <> 0 Then
                If gVerbose Then WScript.Echo "[SONDA JSON] Error obteniendo Left del control " & i & ": " & Err.Description
                Err.Clear
            End If
            
            json = json & "      ""top"": " & ctrl.Top & "," & vbCrLf
            If Err.Number <> 0 Then
                If gVerbose Then WScript.Echo "[SONDA JSON] Error obteniendo Top del control " & i & ": " & Err.Description
                Err.Clear
            End If
            
            json = json & "      ""width"": " & ctrl.Width & "," & vbCrLf
            If Err.Number <> 0 Then
                If gVerbose Then WScript.Echo "[SONDA JSON] Error obteniendo Width del control " & i & ": " & Err.Description
                Err.Clear
            End If
            
            json = json & "      ""height"": " & ctrl.Height & vbCrLf
            If Err.Number <> 0 Then
                If gVerbose Then WScript.Echo "[SONDA JSON] Error obteniendo Height del control " & i & ": " & Err.Description
                Err.Clear
            End If
            
            ' SONDA JSON 4: Agregar propiedades específicas según el tipo de control
            If gVerbose Then WScript.Echo "[SONDA JSON] Agregando propiedades especificas para tipo " & ctrl.ControlType
            
            If ctrl.ControlType = 109 Then ' TextBox
                json = json & "," & vbCrLf & "      ""controlSource"": """ & EscapeJsonString(ctrl.ControlSource) & """"
                If Err.Number <> 0 Then
                    If gVerbose Then WScript.Echo "[SONDA JSON] Error obteniendo ControlSource del TextBox: " & Err.Description
                    Err.Clear
                End If
            ElseIf ctrl.ControlType = 104 Then ' Label
                json = json & "," & vbCrLf & "      ""caption"": """ & EscapeJsonString(ctrl.Caption) & """"
                If Err.Number <> 0 Then
                    If gVerbose Then WScript.Echo "[SONDA JSON] Error obteniendo Caption del Label: " & Err.Description
                    Err.Clear
                End If
            ElseIf ctrl.ControlType = 105 Then ' CommandButton
                json = json & "," & vbCrLf & "      ""caption"": """ & EscapeJsonString(ctrl.Caption) & """"
                If Err.Number <> 0 Then
                    If gVerbose Then WScript.Echo "[SONDA JSON] Error obteniendo Caption del CommandButton: " & Err.Description
                    Err.Clear
                End If
            End If
            
            json = json & vbCrLf & "    }"
        End If
    Next
    
    json = json & vbCrLf & "  ]" & vbCrLf
    json = json & "}" & vbCrLf
    
    ' SONDA JSON 5: Finalizar generación
    If gVerbose Then WScript.Echo "[SONDA JSON] JSON generado exitosamente, longitud: " & Len(json) & " caracteres"
    
    On Error GoTo 0
    
    GenerateFormJson = json
End Function

Function GetControlTypeName(controlType)
    Select Case controlType
        Case 104: GetControlTypeName = "Label"
        Case 105: GetControlTypeName = "CommandButton"
        Case 106: GetControlTypeName = "OptionGroup"
        Case 107: GetControlTypeName = "OptionButton"
        Case 108: GetControlTypeName = "CheckBox"
        Case 109: GetControlTypeName = "TextBox"
        Case 110: GetControlTypeName = "ListBox"
        Case 111: GetControlTypeName = "ComboBox"
        Case 112: GetControlTypeName = "Subform"
        Case 103: GetControlTypeName = "Rectangle"
        Case 102: GetControlTypeName = "Line"
        Case 119: GetControlTypeName = "Image"
        Case Else: GetControlTypeName = "Unknown(" & controlType & ")"
    End Select
End Function

Function EscapeJsonString(str)
    If IsNull(str) Then
        EscapeJsonString = ""
        Exit Function
    End If
    
    Dim result
    result = CStr(str)
    result = Replace(result, "\", "\\")
    result = Replace(result, """", "\""")
    result = Replace(result, vbCrLf, "\n")
    result = Replace(result, vbCr, "\r")
    result = Replace(result, vbLf, "\n")
    result = Replace(result, vbTab, "\t")
    
    EscapeJsonString = result
End Function

' Ejecutar función principal
Main()
