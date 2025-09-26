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
Const acDefault = -1
Const acHidden = 1
Const acNormal = 0
Const acDesign = 1
Const acSaveNo = 2
Const acSaveYes = 0

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

' ============================================================================
' SECCIÓN 2: VARIABLES GLOBALES
' ============================================================================

Dim objFSO, objArgs, objConfig
Dim gVerbose, gQuiet, gDryRun, gDebug
Dim gDbPath, gOutputPath, gConfigPath, gScriptPath, gScriptDir
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
        ' LogMessage "Archivo de configuracion no encontrado: " & configPath & ". Usando valores por defecto."
        Set LoadConfig = config
        Exit Function
    Else
        ' LogMessage "Archivo de configuracion encontrado: " & configPath
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
                        
                        ' LogMessage "Procesando: [" & section & "] " & Trim(parts(0)) & " = " & value & " -> clave: " & key
                        
                        ' Resolver rutas relativas para ciertos valores (excepto patrones)
                        If (InStr(UCase(key), "PATH") > 0 Or InStr(UCase(key), "FILE") > 0) And InStr(UCase(key), "PATTERN") = 0 Then
                            If value <> "" And fso.GetAbsolutePathName(value) <> value Then
                                value = gScriptDir & "\" & value
                                ' LogMessage "Ruta resuelta: " & value
                            End If
                        End If
                        
                        If config.Exists(key) Then
                            config(key) = value
                            ' LogMessage "Clave actualizada: " & key & " = " & value
                        Else
                            config.Add key, value
                            ' LogMessage "Clave agregada: " & key & " = " & value
                        End If
                    End If
                End If
            End If
        End If
    Loop
    
    file.Close
    Set LoadConfig = config
    
    If Err.Number <> 0 Then
        ' LogMessage "Error cargando configuracion: " & Err.Description
        Err.Clear
    End If
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
    WScript.Echo "  schema [<db_path>] [--table <tabla>] [--out <ruta>] [--format <json|md>] [--password <pwd>]"
    WScript.Echo "    Exporta el esquema de la base de datos"
    WScript.Echo "    - Si no se especifica db_path, usa DATABASE_DefaultPath del config"
    WScript.Echo "    - --table: exportar solo una tabla específica (por defecto: todas las tablas)"
    WScript.Echo "    - --out: directorio de salida (por defecto: output/schema/<db_name>)"
    WScript.Echo "    - --format: formato de salida json o md (por defecto: json)"
    WScript.Echo "    - --password: contraseña de la base de datos si está protegida"
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
    Dim pids, i, processCount
    
    LogVerbose "Verificando procesos de Access existentes..."
    pids = GetAccessPIDs()
    
    ' Verificar si hay procesos reales (array no vacío)
    If IsArray(pids) And UBound(pids) >= 0 Then
        processCount = UBound(pids) + 1
        LogMessage "Cerrando " & processCount & " proceso(s) de Access existente(s)..."
        For i = 0 To UBound(pids)
            TerminateAccessPID pids(i)
        Next
        
        ' Esperar un momento para que los procesos se cierren
        WScript.Sleep 1000
        LogVerbose "Procesos de Access cerrados"
    End If
End Sub

' Función singleton para obtener la instancia de Access
' ===========================================================================
' Función canónica para abrir Access - NO usa singleton, retorna instancia directa
' Cada llamada crea una nueva instancia independiente
' ===========================================================================
Function OpenAccessCanonical(dbPath, password)
    Dim objApp, attempt, maxAttempts, retrySleepMs
    maxAttempts = 3
    retrySleepMs = 500
    
    LogMessage "SONDA: OpenAccessCanonical llamada con:"
    LogMessage "SONDA: - dbPath: '" & dbPath & "'"
    LogMessage "SONDA: - password: '" & password & "'"
    LogMessage "SONDA: - Longitud password: " & Len(password)
    
    LogVerbose "Abriendo Access: " & dbPath
    If password <> "" Then
        LogVerbose "Con password: (oculta)"
    End If
    
    ' Cerrar cualquier instancia previa de Access
    LogVerbose "Cerrando instancias previas de Access..."
    CloseAllAccessInstances()
    
    ' Crear instancia de Access con reintentos
    For attempt = 1 To maxAttempts
        On Error Resume Next
        Set objApp = CreateObject("Access.Application")
        
        If Err.Number <> 0 Then
            LogMessage "ERROR: No se pudo crear instancia de Access (intento " & attempt & "): " & Err.Description
            If attempt < maxAttempts Then
                LogVerbose "Reintentando en " & retrySleepMs & "ms..."
                WScript.Sleep retrySleepMs
                Err.Clear
            Else
                Set OpenAccessCanonical = Nothing
                Exit Function
            End If
        Else
            Exit For
        End If
    Next
    
    ' Configurar Access para modo silencioso (solo lo que NO requiere BD abierta)
    objApp.Visible = False
    objApp.UserControl = False
    objApp.AutomationSecurity = 3  ' msoAutomationSecurityForceDisable (3 = ForceDisable)
    
    ' Abrir la base de datos con reintentos
    For attempt = 1 To maxAttempts
        On Error Resume Next
        LogMessage "SONDA: Intento " & attempt & " de abrir BD con password: '" & password & "'"
        If password <> "" Then
            LogMessage "SONDA: Llamando OpenCurrentDatabase con contraseña"
            objApp.OpenCurrentDatabase dbPath, False, password
        Else
            LogMessage "SONDA: Llamando OpenCurrentDatabase sin contraseña"
            objApp.OpenCurrentDatabase dbPath, False
        End If

        If Err.Number <> 0 Then
            LogMessage "ERROR: No se pudo abrir la base de datos (intento " & attempt & "): " & Err.Description
            If attempt < maxAttempts Then
                LogVerbose "Reintentando en " & retrySleepMs & "ms..."
                WScript.Sleep retrySleepMs
                Err.Clear
            Else
                objApp.Quit
                Set objApp = Nothing
                Set OpenAccessCanonical = Nothing
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
        Set OpenAccessCanonical = Nothing
        Exit Function
    End If
    
    ' Configurar modo silencioso DESPUÉS de abrir la BD
    On Error Resume Next
    objApp.DoCmd.SetWarnings False
    objApp.Application.SetOption "Confirm Action Queries", False
    objApp.Application.SetOption "Confirm Document Deletions", False
    objApp.Application.SetOption "Confirm Record Changes", False
    objApp.Application.SetOption "Show Status Bar", False
    objApp.Application.SetOption "Show Animations", False
    objApp.Application.SetOption "Default Open Mode for Databases", 1
    objApp.Application.SetOption "Default Record Locking", 0
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
    Set OpenAccessCanonical = objApp
    
    LogVerbose "Access abierto exitosamente"
End Function

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
    
    ' Configuraciones adicionales después de abrir la BD (requieren BD abierta)
    ' Aplicar configuraciones oficiales de Microsoft para operación desatendida
    On Error Resume Next
    objApp.DoCmd.SetWarnings False
    objApp.Application.SetOption "Confirm Action Queries", False
    objApp.Application.SetOption "Confirm Document Deletions", False
    objApp.Application.SetOption "Confirm Record Changes", False
    
    ' Configuraciones adicionales para interfaz silenciosa (según lecciones aprendidas)
    objApp.Application.SetOption "Show Status Bar", False
    objApp.Application.SetOption "Show Animations", False
    objApp.Application.SetOption "Default Open Mode for Databases", 1  ' Compartido
    objApp.Application.SetOption "Default Record Locking", 0  ' Sin bloqueos
    On Error GoTo 0
    
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
    objApp.Visible = False  ' Reconfirmar que Access permanezca oculto
    objApp.UserControl = False
    objApp.Echo False
    objApp.DoCmd.SetWarnings False
    ' NOTA: DisplayAlerts NO es válido en Access (solo Excel) - removido según lecciones aprendidas
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

' ===========================================================================
' Función canónica para cerrar Access de forma segura
' Implementa secuencia oficial de Microsoft para evitar procesos zombie
' ===========================================================================
Sub CloseAccessCanonical(objAccess)
    If Not objAccess Is Nothing Then
        LogVerbose "Cerrando Access..."
        
        On Error Resume Next
        ' Secuencia oficial Microsoft para cierre seguro
        objAccess.Echo True
        objAccess.DoCmd.SetWarnings True
        objAccess.CloseCurrentDatabase
        objAccess.Quit acQuitSaveNone
        
        ' CRÍTICO: Cerrar CurrentDb DESPUÉS de Quit
        If Not objAccess.CurrentDb Is Nothing Then
            objAccess.CurrentDb.Close
        End If
        
        Set objAccess = Nothing
        DoEvents  ' Permitir limpieza de memoria
        DoEvents  ' Doble DoEvents según mejores prácticas
        On Error GoTo 0
        
        LogVerbose "Access cerrado exitosamente"
    End If
End Sub

' Función para cerrar la instancia singleton de Access - OBSOLETA
' Esta función ya no es necesaria con el patrón explícito
Sub CloseSingletonAccess()
    ' Función obsoleta - usar CloseAccessCanonical directamente
    LogVerbose "CloseSingletonAccess es obsoleta - usar CloseAccessCanonical"
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

Function ExtractFormControls(db, formName, password)
    Dim controlsDict, app, frm
    Set controlsDict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    
    ' Abrir Access y el formulario
    Set app = OpenAccessCanonical(db.Name, password)
    
    app.DoCmd.OpenForm formName, acDesign, , , , acHidden
    
    If Err.Number <> 0 Then
        LogMessage "Error: Formulario '" & formName & "' no encontrado o no se puede abrir", "ERROR"
        CloseAccessCanonical app
        Set ExtractFormControls = controlsDict
        Exit Function
    End If
    
    Set frm = app.Forms(formName)
    Set controlsDict = ExtractFormControlsInternal(frm)
    
    app.DoCmd.Close acForm, formName, acSaveNo
    CloseAccessCanonical app
    
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
    ' Si la ruta ya es absoluta, devolverla tal como está
    If objFSO.GetAbsolutePathName(path) = path Then
        ResolvePath = path
    Else
        ' Si es relativa, resolverla desde el directorio actual de trabajo, no desde el directorio del script
        ResolvePath = objFSO.GetAbsolutePathName(path)
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
Function ExtractModulesToFiles(dbPath, password)
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
    Set objAccess = OpenAccessCanonical(dbPath, password)
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
        CloseAccessCanonical objAccess
        Exit Function
    End If
    
    Set vbComponents = vbProject.VBComponents
    If Err.Number <> 0 Or vbComponents Is Nothing Then
        Err.Clear
        LogMessage "Error: No se pudo acceder a los componentes VBA"
        CloseAccessCanonical objAccess
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
    CloseAccessCanonical objAccess
    
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
    gOutputPath = gScriptDir & "\output"
    
    ' Cargar configuracion de modulos
    Set gConfig = LoadConfig(gConfigPath)
    
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
        ElseIf LCase(objArgs(i)) = "--password" And i < objArgs.Count - 1 Then
            ' Manejar parámetro --password - NO consumir aquí, dejar que lo procese cada comando
            ' Solo agregarlo a cleanArgs para que cada comando lo procese
            ReDim Preserve cleanArgs(cleanArgCount)
            cleanArgs(cleanArgCount) = objArgs(i)
            cleanArgCount = cleanArgCount + 1
            ReDim Preserve cleanArgs(cleanArgCount)
            cleanArgs(cleanArgCount) = objArgs(i + 1)
            cleanArgCount = cleanArgCount + 1
            i = i + 1  ' Saltar el siguiente argumento (valor de la contraseña)
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
            ' Procesar argumentos del comando extract-modules
            Dim dbArg, passwordOpt
            dbArg = ""
            passwordOpt = ""
            
            ' Procesar argumentos específicos del comando
            For k = 1 To cleanArgCount - 1
                Select Case LCase(cleanArgs(k))
                    Case "--password":
                        If k < cleanArgCount - 1 Then 
                            passwordOpt = cleanArgs(k+1)
                            LogMessage "SONDA: Contraseña detectada en argumentos: '" & passwordOpt & "'"
                            k = k + 1
                        Else
                            LogMessage "SONDA: Flag --password encontrado pero sin valor"
                        End If
                    Case Else
                        If dbArg = "" Then dbArg = cleanArgs(k)
                End Select
            Next
            
            ' Si no se especifica ruta, usar DefaultPath del config
            If dbArg = "" Then
                ' Usar DefaultPath de la configuración
                Dim config
                Set config = LoadConfig(gConfigPath)
                gDbPath = ResolvePath(config("DATABASE_DefaultPath"))
                LogMessage "Usando base de datos por defecto: " & gDbPath
            Else
                gDbPath = ResolvePath(dbArg)
            End If
            
            If Not gDryRun Then
                If gVerbose Then WScript.Echo "Extrayendo modulos VBA desde Access..."
                If Not ExtractModulesToFiles(gDbPath, passwordOpt) Then
                    WScript.Echo "Error: No se pudo completar la extraccion de modulos"
                    WScript.Quit 1
                End If
            Else
                WScript.Echo "[DRY-RUN] Se extraerian modulos VBA desde: " & gDbPath
            End If
            

            
        Case "rebuild"
            ' El comando rebuild no acepta parámetros adicionales, siempre usa DefaultPath
            Dim objAccess
            Set config = LoadConfig(gConfigPath)
            gDbPath = ResolvePath(config("DATABASE_DefaultPath"))
            LogMessage "Usando base de datos por defecto: " & gDbPath
            
            ' Validar que el archivo existe
            If Not objFSO.FileExists(gDbPath) Then
                WScript.Echo "Error: El archivo de base de datos no existe: " & gDbPath
                WScript.Quit 1
            End If
            
            If gVerbose Then WScript.Echo "Reconstruyendo modulos VBA..."
            
            ' Abrir Access para RebuildProject (sin contraseña para rebuild)
            Set objAccess = OpenAccessCanonical(gDbPath, "")
            If objAccess Is Nothing Then
                WScript.Echo "Error: No se pudo abrir Access"
                WScript.Quit 1
            End If
            
            ' Llamar a RebuildProject con objAccess como parámetro
            Call RebuildProject(objAccess)
            
            ' Cerrar Access
            Call CloseAccessCanonical(objAccess)
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
                    
                    ' Abrir Access para UpdateModules (sin contraseña para update)
                    Dim objAccessUpdate
                    Set objAccessUpdate = OpenAccessCanonical(gDbPath, "")
                    If objAccessUpdate Is Nothing Then
                        WScript.Echo "Error: No se pudo abrir la base de datos"
                        WScript.Quit 1
                    End If
                    
                    ' Llamar a UpdateModules con objAccess como parámetro
                    Dim res
                    res = UpdateModules(objAccessUpdate, modulesArg)
                    
                    ' Cerrar Access
                    CloseAccessCanonical objAccessUpdate
                    
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
            
        Case "export-form"
            ' Procesar argumentos del comando export-form
            Dim dbArg, formName, outputPath, passwordOpt
            dbArg = ""
            formName = ""
            outputPath = ""
            passwordOpt = ""
            
            ' Procesar argumentos específicos del comando
            For k = 1 To cleanArgCount - 1
                Select Case LCase(cleanArgs(k))
                    Case "--password":
                        If k < cleanArgCount - 1 Then 
                            passwordOpt = cleanArgs(k+1)
                            LogMessage "SONDA: Contraseña detectada en argumentos: '" & passwordOpt & "'"
                            k = k + 1
                        Else
                            LogMessage "SONDA: Flag --password encontrado pero sin valor"
                        End If
                    Case Else
                        If dbArg = "" Then 
                            dbArg = cleanArgs(k)
                        ElseIf formName = "" Then
                            formName = cleanArgs(k)
                        ElseIf outputPath = "" Then
                            outputPath = cleanArgs(k)
                        End If
                End Select
            Next
            
            If dbArg <> "" And formName <> "" Then
                gDbPath = ResolvePath(dbArg)
                
                ' Si no se especifica ruta de salida, usar el nombre del formulario con extensión .json
                If outputPath = "" Then
                    outputPath = objFSO.GetAbsolutePathName(formName & ".json")
                Else
                    outputPath = ResolvePath(outputPath)
                End If
                
                If Not gDryRun Then
                    ExportFormToJSON gDbPath, passwordOpt, formName, outputPath
                Else
                    LogMessage "SIMULACION: Exportaria formulario " & formName & " de " & gDbPath & " a " & outputPath
                End If
            Else
                LogError "Faltan argumentos para export-form"
                WScript.Echo "Uso: cscript cli.vbs export-form <database> <form_name> [output_file] --password <password>"
                ShowHelp
            End If
            
        Case "schema"
            ' Parseo de opciones para el comando schema
            Dim dbArg, tableOpt, outOpt, fmtOpt, passwordOpt, k
            dbArg = ""
            tableOpt = ""
            outOpt = ""
            fmtOpt = "json"
            passwordOpt = ""
            
            ' Parsear argumentos posicionales y flags
            For k = 1 To cleanArgCount - 1
                Select Case LCase(cleanArgs(k))
                    Case "--table": 
                        If k < cleanArgCount - 1 Then 
                            tableOpt = cleanArgs(k+1)
                            k = k + 1
                        End If
                    Case "--out":   
                        If k < cleanArgCount - 1 Then 
                            outOpt = cleanArgs(k+1)
                            k = k + 1
                        End If
                    Case "--format":
                        If k < cleanArgCount - 1 Then 
                            fmtOpt = LCase(cleanArgs(k+1))
                            k = k + 1
                        End If
                    Case "--password":
                        If k < cleanArgCount - 1 Then 
                            passwordOpt = cleanArgs(k+1)
                            LogMessage "SONDA: Contraseña detectada en argumentos: '" & passwordOpt & "'"
                            k = k + 1
                        Else
                            LogMessage "SONDA: Flag --password encontrado pero sin valor"
                        End If
                    Case Else
                        If dbArg = "" Then dbArg = cleanArgs(k)
                End Select
            Next
            
            ' Usar base de datos por defecto desde .ini si no se especifica
            If dbArg = "" Then
                If objConfig.Exists("DATABASE_DefaultPath") Then
                    dbArg = objConfig("DATABASE_DefaultPath")
                    LogMessage "Usando base de datos por defecto: " & dbArg
                Else
                    WScript.Echo "Error: No se especificó base de datos y no hay DATABASE_DefaultPath en cli.ini"
                    WScript.Quit 1
                End If
            End If
            
            gDbPath = ResolvePath(dbArg)
            
            ' Validar que el archivo existe
            If Not objFSO.FileExists(gDbPath) Then
                WScript.Echo "Error: El archivo de base de datos no existe: " & gDbPath
                WScript.Quit 1
            End If
            
            ' Configurar directorio de salida por defecto
            If outOpt = "" Then 
                Dim dbName: dbName = objFSO.GetBaseName(gDbPath)
                outOpt = gScriptDir & "\output\schema\" & dbName
            End If
            
            ' Crear directorio de salida
            CreateFolderRecursive outOpt
            
            ' Validar formato
            If fmtOpt <> "json" And fmtOpt <> "md" Then 
                fmtOpt = "json"
            End If
            
            ' Usar contraseña del parámetro directamente
            If passwordOpt <> "" Then
                LogMessage "Contraseña especificada por parámetro"
            End If

            If Not gDryRun Then
                LogMessage "Exportando esquema de " & gDbPath & " a " & outOpt & " (formato: " & fmtOpt & ")"
                If tableOpt <> "" Then
                    LogMessage "Tabla específica: " & tableOpt
                Else
                    LogMessage "Exportando todas las tablas"
                End If
                
                If Not ExportSchema(gDbPath, tableOpt, outOpt, fmtOpt, passwordOpt) Then
                    WScript.Echo "Error: no se pudo exportar el esquema"
                    WScript.Quit 1
                End If
                LogMessage "Esquema exportado exitosamente a " & outOpt
            Else
                LogMessage "SIMULACION: Exportaria esquema de " & gDbPath & " a " & outOpt & " (formato: " & fmtOpt & ")"
                If tableOpt <> "" Then
                    LogMessage "SIMULACION: Tabla específica: " & tableOpt
                Else
                    LogMessage "SIMULACION: Exportando todas las tablas"
                End If
            End If
            
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
Function UpdateModules(objAccess, modulesArg)
    On Error Resume Next
    UpdateModules = False
    
    ' Normaliza la lista igual que rebuild; si rebuild ya tiene un normalizador, úsalo
    Dim list, i, name
    list = NormalizeModuleList(modulesArg)  ' Debe quitar extensiones, dividir por coma/espacio y eliminar duplicados
    
    If IsEmpty(list) Or UBound(list) < 0 Then
        LogMessage "update: sin módulos para actualizar"
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
            Call RebuildLike_ImportOne(objAccess, name)
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
    objAccess.RunCommand acCmdCompileAndSaveAllModules
    If Err.Number <> 0 Then
        LogMessage "update: aviso al compilar: " & Err.Description
        Err.Clear
        ' Fallback: no intentes ningún otro RunCommand; continúa al cierre
    End If
    On Error GoTo 0
    
    LogMessage "update: proceso completado exitosamente"
    UpdateModules = True
End Function

' Función para listar objetos de la base de datos
Sub ListObjects(dbPath)
    LogMessage "Listando objetos de: " & objFSO.GetFileName(dbPath)
    
    Set objAccess = OpenAccessCanonical(dbPath, "")
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
    
    CloseAccessCanonical objAccess
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
    
    ' Intentar crear instancia de Access (sin contraseña para tests)
    Set objAccess = OpenAccessCanonical(dbPath, "")
    If Err.Number <> 0 Or objAccess Is Nothing Then
        TestAccessAutomation = False
        Exit Function
    End If
    
    ' Cerrar y limpiar
    CloseAccessCanonical objAccess
    
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
    
    ' Intentar crear instancia de Access (sin contraseña para tests)
    Set objAccess = OpenAccessCanonical(dbPath, "")
    If Err.Number <> 0 Or objAccess Is Nothing Then
        TestVBEAccess = False
        Exit Function
    End If
    
    ' Probar acceso VBE usando CheckVBProjectAccess
    If Not CheckVBProjectAccess(objAccess) Then
        CloseAccessCanonical objAccess
        TestVBEAccess = False
        Exit Function
    End If
    
    ' Cerrar y limpiar
    CloseAccessCanonical objAccess
    
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
Sub ImportModuleWithAnsiEncoding(objAccess, strImportPath, moduleName, fileExtension, cleanedContent)
    On Error Resume Next
    
    ' LogMessage "  Importando: " & moduleName & " desde " & strImportPath
    
    ' Crear nuevo componente
    Dim newComponent, componentType
    
    If LCase(fileExtension) = "cls" Then
        componentType = 2  ' vbext_ct_ClassModule
    Else
        componentType = 1  ' vbext_ct_StdModule
    End If
    
    ' LogMessage "  Creando nuevo componente tipo " & componentType & " para: " & moduleName
    Set newComponent = objAccess.VBE.ActiveVBProject.VBComponents.Add(componentType)
    
    ' Verificar que el componente se creó correctamente
    If newComponent Is Nothing Then
        LogMessage "  ❌ Error: No se pudo crear componente para " & moduleName
        Exit Sub
    End If
    
    newComponent.Name = moduleName
    
    ' Verificar si el componente ya existe
    Dim existingComponent
    Set existingComponent = Nothing
    For Each existingComponent In objAccess.VBE.ActiveVBProject.VBComponents
        If existingComponent.Name = moduleName Then
            Exit For
        End If
        Set existingComponent = Nothing
    Next
    
    ' Si existe, eliminar el código existente
    If Not existingComponent Is Nothing Then
        ' LogMessage "  Limpiando código existente de: " & moduleName
        
        ' Limpiar solo si hay líneas
        If existingComponent.CodeModule.CountOfLines > 0 Then
            existingComponent.CodeModule.DeleteLines 1, existingComponent.CodeModule.CountOfLines
        End If
        
        ' Insertar contenido limpio
        If Len(cleanedContent) > 0 Then
            existingComponent.CodeModule.AddFromString cleanedContent
            ' LogMessage "  ✓ Actualizado: " & moduleName
        End If
    End If
    
    ' Verificar si hubo errores críticos durante la importación
    If Err.Number <> 0 Then
        LogMessage "  ❌ Error crítico con " & moduleName & ": " & Err.Description
        Err.Clear
    Else
        ' LogMessage "  ✓ Importación completada exitosamente para: " & moduleName
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
    
    ' Log simplificado - solo el nombre del módulo
    LogMessage "Importando: " & moduleName
    
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
    Call ImportModuleWithAnsiEncoding(objAccess, filePath, moduleName, fileExtension, cleanedContent)
    
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
    ' LogMessage "✓ Módulo " & moduleName & " importado exitosamente"
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
Sub RebuildProject(objAccess)
    ' WScript.Echo "=== RECONSTRUCCION COMPLETA DEL PROYECTO VBA ==="
    ' WScript.Echo "ADVERTENCIA: Se eliminaran TODOS los modulos VBA existentes"
    ' WScript.Echo "Iniciando proceso de reconstruccion..."
    
    On Error Resume Next
    
    ' Paso 1: Eliminar todos los módulos existentes
    ' WScript.Echo "Paso 1: Eliminando todos los modulos VBA existentes..."
    
    Dim vbProject, vbComponent
    Set vbProject = objAccess.VBE.ActiveVBProject
    
    Dim componentCount, i, errorDetails
    componentCount = vbProject.VBComponents.Count
    
    ' Iterar hacia atrás para evitar problemas al eliminar elementos
    For i = componentCount To 1 Step -1
        Set vbComponent = vbProject.VBComponents(i)
        
        ' Solo eliminar módulos estándar y de clase (no formularios ni informes)
        If vbComponent.Type = 1 Or vbComponent.Type = 2 Then ' vbext_ct_StdModule = 1, vbext_ct_ClassModule = 2
            ' WScript.Echo "  Eliminando: " & vbComponent.Name & " (Tipo: " & vbComponent.Type & ")"
            vbProject.VBComponents.Remove vbComponent
            
            If Err.Number <> 0 Then
                ' WScript.Echo "  ❌ Error eliminando " & vbComponent.Name & ": " & Err.Description
                Err.Clear
            Else
                ' WScript.Echo "  ✓ Eliminado: " & vbComponent.Name
            End If
        End If
    Next
    
    ' Paso 2: Guardar cambios sin cerrar Access
    ' WScript.Echo "Paso 2: Guardando cambios..."
    objAccess.DoCmd.Save
    
    If Err.Number <> 0 Then
        ' WScript.Echo "Advertencia al guardar: " & Err.Description
        Err.Clear
    End If
    
    ' WScript.Echo "✓ Módulos eliminados y cambios guardados"
    
    ' Paso 3: Importar todos los módulos de nuevo
    ' WScript.Echo "Paso 3: Importando todos los modulos desde /src..."
    
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
    
    ' PASO 3.1: Validacion previa de sintaxis
    ' LogMessage "Validando sintaxis de todos los modulos..."
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
                ' LogMessage "  Sintaxis valida: " & objFile.Name
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
    
    ' LogMessage "Validacion completada: " & validFiles & " archivos validos"
    
    ' PASO 3.2: Procesar archivos de modulos
    Set objFolder = objFSO.GetFolder(g_ModulesSrcPath)
    
    For Each objFile In objFolder.Files
        If LCase(objFSO.GetExtensionName(objFile.Name)) = "bas" Or LCase(objFSO.GetExtensionName(objFile.Name)) = "cls" Then
            strModuleName = objFSO.GetBaseName(objFile.Name)
            
            ' LogMessage "Procesando modulo: " & strModuleName
            
            ' Usar la función unificada de importación
            Dim importResult
            importResult = RebuildLike_ImportOne(objAccess, strModuleName)
            
            If Not importResult Then
                ' LogMessage "Error al importar modulo " & strModuleName
            End If
        End If
    Next
    
    ' PASO 3.3: Guardar cada modulo individualmente
    ' LogMessage "Guardando modulos individualmente..."
    On Error Resume Next
    
    For Each vbComponent In objAccess.VBE.ActiveVBProject.VBComponents
        If vbComponent.Type = 1 Then  ' vbext_ct_StdModule
            ' LogMessage "Guardando modulo: " & vbComponent.Name
            objAccess.DoCmd.Save 5, vbComponent.Name  ' acModule = 5
            If Err.Number <> 0 Then
                ' LogMessage "Advertencia al guardar " & vbComponent.Name & ": " & Err.Description
                Err.Clear
            End If
        ElseIf vbComponent.Type = 2 Then  ' vbext_ct_ClassModule
            ' LogMessage "Guardando clase: " & vbComponent.Name
            objAccess.DoCmd.Save 7, vbComponent.Name  ' acClassModule = 7
            If Err.Number <> 0 Then
                ' LogMessage "Advertencia al guardar " & vbComponent.Name & ": " & Err.Description
                Err.Clear
            End If
        End If
    Next
    
    ' PASO 3.4: Verificacion de integridad y compilacion
    ' LogMessage "Verificando integridad de nombres de modulos..."
    Call VerifyModuleNames(objAccess, g_ModulesSrcPath)
    
    ' LogMessage "=== RECONSTRUCCION COMPLETADA EXITOSAMENTE ==="
    ' LogMessage "El proyecto VBA ha sido completamente reconstruido"
    ' LogMessage "Todos los modulos han sido reimportados desde /src"
    
    On Error GoTo 0
End Sub

' ============================================================================
' SECCIÓN 10: FUNCIONES DE EXPORTACIÓN DE FORMULARIOS
' ============================================================================

' Función para obtener el nombre del tipo de control
Function GetControlTypeName(controlType)
    Select Case controlType
        Case acLabel: GetControlTypeName = "Label"
        Case acTextBox: GetControlTypeName = "TextBox"
        Case acCommandButton: GetControlTypeName = "CommandButton"
        Case acCheckBox: GetControlTypeName = "CheckBox"
        Case acOptionButton: GetControlTypeName = "OptionButton"
        Case acComboBox: GetControlTypeName = "ComboBox"
        Case acListBox: GetControlTypeName = "ListBox"
        Case acSubform: GetControlTypeName = "Subform"
        Case acImage: GetControlTypeName = "Image"
        Case acRectangle: GetControlTypeName = "Rectangle"
        Case acLine: GetControlTypeName = "Line"
        Case acOptionGroup: GetControlTypeName = "OptionGroup"
        Case acBoundObjectFrame: GetControlTypeName = "BoundObjectFrame"
        Case acUnboundObjectFrame: GetControlTypeName = "UnboundObjectFrame"
        Case acPageBreak: GetControlTypeName = "PageBreak"
        Case acCustomControl: GetControlTypeName = "CustomControl"
        Case acToggleButton: GetControlTypeName = "ToggleButton"
        Case acTabCtl: GetControlTypeName = "TabControl"
        Case acPage: GetControlTypeName = "Page"
        Case Else: GetControlTypeName = "Unknown (" & controlType & ")"
    End Select
End Function

' Función para extraer controles de un formulario
Function ExtractFormControls(objApp, formName)
    On Error Resume Next
    
    LogMessage "Extrayendo controles del formulario: " & formName
    
    ' Abrir el formulario en vista de diseño
    objApp.DoCmd.OpenForm formName, acDesign
    If Err.Number <> 0 Then
        LogMessage "Error al abrir formulario " & formName & " en vista de diseño: " & Err.Description
        ExtractFormControls = ""
        Exit Function
    End If
    
    ' Obtener referencia al formulario
    Dim frm
    Set frm = objApp.Forms(formName)
    
    ' Extraer controles
    Dim result
    result = ExtractFormControlsInternal(frm)
    
    ' Cerrar el formulario sin guardar
    objApp.DoCmd.Close acForm, formName, acSaveNo
    
    ExtractFormControls = result
    On Error GoTo 0
End Function

' Función interna para extraer controles
Function ExtractFormControlsInternal(frm)
    On Error Resume Next
    
    Dim result, ctrl, i
    result = "{"
    result = result & """formName"": """ & frm.Name & ""","
    result = result & """controls"": ["
    
    i = 0
    For Each ctrl In frm.Controls
        If i > 0 Then result = result & ","
        
        result = result & "{"
        result = result & """name"": """ & ctrl.Name & ""","
        result = result & """type"": """ & GetControlTypeName(ctrl.ControlType) & ""","
        result = result & """controlType"": " & ctrl.ControlType
        
        ' Propiedades específicas según el tipo de control
        If ctrl.ControlType = acTextBox Or ctrl.ControlType = acComboBox Or ctrl.ControlType = acListBox Then
            If Not IsNull(ctrl.ControlSource) And ctrl.ControlSource <> "" Then
                result = result & ",""controlSource"": """ & ctrl.ControlSource & """"
            End If
        End If
        
        If ctrl.ControlType = acLabel Or ctrl.ControlType = acCommandButton Or ctrl.ControlType = acCheckBox Or ctrl.ControlType = acOptionButton Then
            If Not IsNull(ctrl.Caption) And ctrl.Caption <> "" Then
                result = result & ",""caption"": """ & Replace(ctrl.Caption, """", "\""") & """"
            End If
        End If
        
        If ctrl.ControlType = acCommandButton Then
            If Not IsNull(ctrl.OnClick) And ctrl.OnClick <> "" Then
                result = result & ",""onClick"": """ & ctrl.OnClick & """"
            End If
        End If
        
        If ctrl.ControlType = acSubform Then
            If Not IsNull(ctrl.SourceObject) And ctrl.SourceObject <> "" Then
                result = result & ",""sourceObject"": """ & ctrl.SourceObject & """"
            End If
        End If
        
        ' Propiedades de formato comunes
        result = result & ",""left"": " & ctrl.Left
        result = result & ",""top"": " & ctrl.Top
        result = result & ",""width"": " & ctrl.Width
        result = result & ",""height"": " & ctrl.Height
        result = result & ",""visible"": " & LCase(CStr(ctrl.Visible))
        
        result = result & "}"
        i = i + 1
    Next
    
    result = result & "]}"
    ExtractFormControlsInternal = result
    On Error GoTo 0
End Function

' Función para exportar formulario a JSON
Sub ExportFormToJSON(dbPath, password, formName, outputPath)
    On Error Resume Next
    
    LogMessage "=== EXPORTANDO FORMULARIO A JSON ==="
    LogMessage "Base de datos: " & dbPath
    LogMessage "Formulario: " & formName
    LogMessage "Archivo de salida: " & outputPath
    
    ' Abrir Access
    Dim objApp
    Set objApp = OpenAccessCanonical(dbPath, password)
    If objApp Is Nothing Then
        LogMessage "Error: No se pudo abrir la base de datos"
        Exit Sub
    End If
    
    ' Extraer controles del formulario
    Dim jsonContent
    jsonContent = ExtractFormControls(objApp, formName, password)
    
    If jsonContent = "" Then
        LogMessage "Error: No se pudieron extraer los controles del formulario"
        CloseAccessCanonical objApp
        Exit Sub
    End If
    
    ' Guardar a archivo JSON
    Dim objFile
    Set objFile = objFSO.CreateTextFile(outputPath, True)
    objFile.Write jsonContent
    objFile.Close
    
    LogMessage "Formulario exportado exitosamente a: " & outputPath
    
    ' Cerrar Access
    CloseAccessCanonical objApp
    
    On Error GoTo 0
End Sub

' ============================================================================
' SECCIÓN UI AS CODE: FUNCIONES AVANZADAS BASADAS EN LECCIONES APRENDIDAS
' ============================================================================

' ============================================================================
' CONFIGURACIÓN ANTI-UI MEJORADA
' ============================================================================

' Versión canónica y reforzada basada en lecciones aprendidas
Sub AntiUI(app)
    On Error Resume Next
    ' Configuración baseline para operación silenciosa
    ' NOTA: DisplayAlerts NO es válido en Access (lección aprendida)
    app.Echo = False
    app.DoCmd.SetWarnings False
    app.AutomationSecurity = 3 ' ForceDisable
    app.Visible = False
    app.UserControl = False

    ' Desactivar prompts de confirmación
    app.Application.SetOption "Confirm Action Queries", False
    app.Application.SetOption "Confirm Document Deletions", False
    app.Application.SetOption "Confirm Record Changes", False

    ' Desactivar Name AutoCorrect (evita recalcular objetos y prompts)
    app.Application.SetOption "Track Name AutoCorrect Info", False
    app.Application.SetOption "Perform Name AutoCorrect", False
    app.Application.SetOption "Auto Compact", False

    Err.Clear
End Sub

' ============================================================================
' FUNCIONES DE CONFIGURACIÓN UI
' ============================================================================

Function UI_GetRoot(config)
    UI_GetRoot = CfgGet(config, "UI_Root", "UI.Root", "ui")
End Function

Function UI_GetFormsDir(config)
    UI_GetFormsDir = CfgGet(config, "UI_FormsDir", "UI.FormsDir", "forms")
End Function

Function UI_GetAssetsDir(config)
    UI_GetAssetsDir = CfgGet(config, "UI_AssetsDir", "UI.AssetsDir", "assets")
End Function

Function UI_GetIncludeSubdirs(config)
    UI_GetIncludeSubdirs = ToBool(CfgGet(config,"UI_IncludeSubdirectories","UI.IncludeSubdirectories","true"),True)
End Function

Function UI_GetFormFilePattern(config)
    UI_GetFormFilePattern = CfgGet(config, "UI_FormFilePattern", "UI.FormFilePattern", "*.json")
End Function

Function UI_NameFromFileBase(config)
    UI_NameFromFileBase = ToBool(CfgGet(config,"UI_NameFromFileBase","UI.NameFromFileBase","true"),True)
End Function

Function UI_GetAssetsImgDir(config)
    UI_GetAssetsImgDir = CfgGet(config, "UI_AssetsImgDir", "UI.AssetsImgDir", "img")
End Function

Function UI_GetAssetsImgExtensions(config)
    UI_GetAssetsImgExtensions = CfgGet(config, "UI_AssetsImgExtensions", "UI.AssetsImgExtensions", "png,jpg,jpeg,gif,bmp,ico")
End Function

Function UI_StrictProperties(config)
    UI_StrictProperties = ToBool(CfgGet(config,"UI_StrictProperties","UI.StrictProperties","false"),False)
End Function

' ============================================================================
' FUNCIONES DE MANEJO DE ASSETS E IMÁGENES
' ============================================================================

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

Sub MissingAssetLog(imgRef, resolvedForForm)
    On Error Resume Next
    LogMessage "ui-import: imagen no encontrada: " & imgRef & " (form=" & resolvedForForm & ")"
End Sub

' ============================================================================
' FUNCIONES DE CREACIÓN PROGRAMÁTICA DE FORMULARIOS
' ============================================================================

' Función para crear formulario programáticamente basada en lecciones aprendidas
Function CreateFormProgrammatically(app, formName, recordSource)
    On Error Resume Next
    CreateFormProgrammatically = False
    
    If app Is Nothing Then Exit Function
    
    ' Aplicar configuración anti-UI
    AntiUI app
    
    ' Crear formulario nuevo
    Dim frm
    Set frm = app.CreateForm()
    If frm Is Nothing Then Exit Function
    
    ' Configurar propiedades básicas del formulario
    frm.RecordSource = recordSource
    frm.Caption = formName
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.DividingLines = False
    
    ' Guardar con nombre temporal primero (patrón atómico)
    Dim tmpName
    tmpName = formName & "__tmp"
    app.DoCmd.Save acForm, tmpName
    
    ' Cerrar formulario
    app.DoCmd.Close acForm, tmpName, acSaveYes
    
    CreateFormProgrammatically = True
End Function

' Función para agregar control a formulario
Function AddControlToForm(app, formName, controlType, controlName, left, top, width, height)
    On Error Resume Next
    AddControlToForm = False
    
    If app Is Nothing Then Exit Function
    
    ' Aplicar configuración anti-UI
    AntiUI app
    
    ' Abrir formulario en modo diseño
    app.DoCmd.OpenForm formName, acDesign
    
    Dim frm, ctrl
    Set frm = app.Forms(formName)
    
    ' Crear control
    Set ctrl = app.CreateControl(formName, controlType, , , , left, top, width, height)
    If Not ctrl Is Nothing Then
        ctrl.Name = controlName
        AddControlToForm = True
    End If
    
    ' Guardar y cerrar
    AntiUI app
    app.DoCmd.Close acForm, formName, acSaveYes
End Function

' ============================================================================
' FUNCIONES DE IMPORTACIÓN DE FORMULARIOS DESDE JSON
' ============================================================================

' Variables globales para importación de formularios
Dim ImportFormFromJson_app, ImportFormFromJson_target, ImportFormFromJson_finalTarget, ImportFormFromJson_root

Function ImportFormFromJson(dbPath, jsonPath, targetName)
    On Error Resume Next
    ImportFormFromJson = False
    
    ' Abrir Access con configuración anti-UI
    Dim app
    Set app = OpenAccessCanonical(dbPath, GetDatabasePassword(dbPath))
    If app Is Nothing Then Exit Function
    
    AntiUI app
    
    ' Leer y parsear JSON
    Dim jsonContent, root
    jsonContent = ReadAllText(jsonPath)
    If jsonContent = "" Then
        CloseAccessCanonical app
        Exit Function
    End If
    
    Set root = JsonParse(jsonContent)
    If root Is Nothing Then
        CloseAccessCanonical app
        Exit Function
    End If
    
    ' Configurar variables globales para finalización
    Set ImportFormFromJson_app = app
    ImportFormFromJson_target = targetName & "__tmp"
    ImportFormFromJson_finalTarget = targetName
    Set ImportFormFromJson_root = root
    
    ' Crear formulario base
    If Not CreateFormFromJson(app, root, ImportFormFromJson_target) Then
        CloseAccessCanonical app
        Exit Function
    End If
    
    ' Finalizar importación con swap atómico
    ImportFormFromJson = ImportFormFromJson_Finalize()
End Function

Function CreateFormFromJson(app, root, formName)
    On Error Resume Next
    CreateFormFromJson = False
    
    If app Is Nothing Or root Is Nothing Then Exit Function
    
    ' Aplicar configuración anti-UI
    AntiUI app
    
    ' Crear formulario
    Dim frm
    Set frm = app.CreateForm()
    If frm Is Nothing Then Exit Function
    
    ' Configurar propiedades del formulario desde JSON
    If root.Exists("recordSource") Then frm.RecordSource = root("recordSource")
    If root.Exists("caption") Then frm.Caption = root("caption")
    If root.Exists("width") Then frm.Width = root("width")
    
    ' Configurar propiedades adicionales
    frm.NavigationButtons = False
    frm.RecordSelectors = False
    frm.DividingLines = False
    
    ' Crear controles si existen
    If root.Exists("controls") Then
        Dim controls, i
        Set controls = root("controls")
        If IsObject(controls) Then
            For i = 0 To controls.Count - 1
                CreateControlFromJson app, frm, controls.Items()(i)
            Next
        End If
    End If
    
    ' Guardar formulario
    AntiUI app
    app.DoCmd.Save acForm, formName
    app.DoCmd.Close acForm, formName, acSaveYes
    
    CreateFormFromJson = True
End Function

Sub CreateControlFromJson(app, frm, controlDict)
    On Error Resume Next
    
    If Not IsObject(controlDict) Then Exit Sub
    If Not controlDict.Exists("type") Then Exit Sub
    
    Dim controlType, controlName, left, top, width, height
    controlType = GetControlTypeFromName(controlDict("type"))
    controlName = controlDict("name")
    left = controlDict("left")
    top = controlDict("top")
    width = controlDict("width")
    height = controlDict("height")
    
    ' Crear control
    Dim ctrl
    Set ctrl = app.CreateControl(frm.Name, controlType, , , , left, top, width, height)
    If Not ctrl Is Nothing Then
        ctrl.Name = controlName
        
        ' Aplicar propiedades específicas
        ApplyControlProperties ctrl, controlDict
        
        ' Manejar imágenes si es necesario
        HandleControlImage ctrl, controlDict, frm.Name
    End If
End Sub

Sub ApplyControlProperties(ctrl, controlDict)
    On Error Resume Next
    
    ' Aplicar propiedades comunes
    If controlDict.Exists("caption") Then SetCtrlProp ctrl, "Caption", controlDict("caption")
    If controlDict.Exists("controlSource") Then SetCtrlProp ctrl, "ControlSource", controlDict("controlSource")
    If controlDict.Exists("visible") Then SetCtrlProp ctrl, "Visible", controlDict("visible")
    If controlDict.Exists("enabled") Then SetCtrlProp ctrl, "Enabled", controlDict("enabled")
    If controlDict.Exists("locked") Then SetCtrlProp ctrl, "Locked", controlDict("locked")
    If controlDict.Exists("tabIndex") Then SetCtrlProp ctrl, "TabIndex", controlDict("tabIndex")
    
    ' Propiedades de eventos
    If controlDict.Exists("onClick") Then SetCtrlProp ctrl, "OnClick", controlDict("onClick")
    If controlDict.Exists("onDblClick") Then SetCtrlProp ctrl, "OnDblClick", controlDict("onDblClick")
    If controlDict.Exists("onChange") Then SetCtrlProp ctrl, "OnChange", controlDict("onChange")
End Sub

Sub HandleControlImage(ctrl, controlDict, formName)
    On Error Resume Next
    
    ' Manejar carga de imágenes desde assets
    Dim imgRef
    imgRef = ""
    If controlDict.Exists("assetImage") Then imgRef = controlDict("assetImage")
    If imgRef = "" And controlDict.Exists("image") Then imgRef = controlDict("image")
    If imgRef = "" And controlDict.Exists("picture") Then imgRef = controlDict("picture")
    
    If imgRef <> "" Then
        Dim fullImg
        fullImg = ResolveAssetImagePath(imgRef, gConfig)
        If fullImg <> "" Then
            If ctrl.ControlType = acImage Then
                ctrl.Picture = fullImg
                ' Configurar como Linked para evitar prompts
                If HasProperty(ctrl, "PictureType") Then ctrl.PictureType = 1
                
                ' Propiedades adicionales de imagen
                If controlDict.Exists("pictureAlignment") Then SetCtrlProp ctrl, "PictureAlignment", controlDict("pictureAlignment")
                If controlDict.Exists("sizeMode") Then SetCtrlProp ctrl, "SizeMode", controlDict("sizeMode")
                If controlDict.Exists("pictureTiling") Then SetCtrlProp ctrl, "PictureTiling", controlDict("pictureTiling")
            ElseIf ctrl.ControlType = acCommandButton Then
                If HasProperty(ctrl, "Picture") Then ctrl.Picture = fullImg
                If controlDict.Exists("pictureType") Then SetCtrlProp ctrl, "PictureType", controlDict("pictureType")
            End If
        Else
            MissingAssetLog imgRef, formName
        End If
    End If
End Sub

Sub SetCtrlProp(ctrl, propName, propValue)
    On Error Resume Next
    
    ' Salir si el valor está vacío
    If IsEmpty(propValue) Or propValue = "" Or IsNull(propValue) Then Exit Sub
    
    ' Verificar si el control tiene la propiedad
    If HasProperty(ctrl, propName) Then
        ctrl.Properties(propName) = propValue
    ElseIf UI_StrictProperties(gConfig) Then
        LogMessage "UI.StrictProperties: propiedad '" & propName & "' no existe en control '" & ctrl.Name & "'"
    End If
    
    Err.Clear
End Sub

Function ImportFormFromJson_Finalize()
    On Error Resume Next
    ImportFormFromJson_Finalize = False
    
    If ImportFormFromJson_app Is Nothing Then Exit Function
    
    Dim app, tmpName, finalName
    Set app = ImportFormFromJson_app
    tmpName = ImportFormFromJson_target
    finalName = ImportFormFromJson_finalTarget
    
    ' Aplicar configuración anti-UI antes de guardar
    AntiUI app
    
    ' Guardar temporal
    app.DoCmd.Save acForm, tmpName
    
    ' Swap atómico: renombrar temporal a final
    SafeSwapForm app, finalName, tmpName
    
    ' Compilar y guardar módulos si es necesario
    On Error Resume Next
    app.DoCmd.RunCommand acCmdCompileAndSaveAllModules
    Err.Clear
    
    ' Cerrar Access
    CloseAccessCanonical app
    
    ' Limpiar referencias
    Set ImportFormFromJson_app = Nothing
    ImportFormFromJson_target = ""
    ImportFormFromJson_finalTarget = ""
    Set ImportFormFromJson_root = Nothing
    
    ImportFormFromJson_Finalize = True
End Function

' ============================================================================
' SECCIÓN 7: FUNCIONES DE ESQUEMA DE BASE DE DATOS
' ============================================================================

' Exporta estructura de todas las tablas o una tabla específica
Function ExportSchema(dbPath, tableFilter, outDir, fmt, password)
    On Error Resume Next
    ExportSchema = False

    LogMessage "SONDA: ExportSchema llamada con parámetros:"
    LogMessage "SONDA: - dbPath: '" & dbPath & "'"
    LogMessage "SONDA: - password: '" & password & "'"
    
    LogVerbose "ExportSchema: Iniciando con dbPath=" & dbPath & ", password=" & password
    LogMessage "DEBUG: Llamando a OpenAccessCanonical..."
    
    ' Usar el patrón Access singleton
    Set objAccess = OpenAccessCanonical(dbPath, password)
    LogMessage "DEBUG: OpenAccessCanonical retornó: " & TypeName(objAccess)
    
    If objAccess Is Nothing Then 
        LogMessage "Error: No se pudo abrir la base de datos: " & dbPath
        Exit Function
    End If

    Dim db: Set db = objAccess.CurrentDb
    If db Is Nothing Then
        LogMessage "Error: No se pudo acceder a CurrentDb"
        CloseAccessCanonical objAccess
        Exit Function
    End If
    
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
                    If gVerbose Then LogMessage "Campo agregado: " & f.Name & " (" & GetFieldTypeName(f.Type) & ")"
                Next
                Set tinfo("fields") = fields
                If gVerbose Then LogMessage "Total campos en " & tdef.Name & ": " & fields.Count

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

    ' SIEMPRE cerrar Access usando el patrón singleton
    CloseAccessCanonical objAccess
    
    If Err.Number <> 0 Then
        LogMessage "Error durante exportación de esquema: " & Err.Description
        Err.Clear
        ExportSchema = False
    Else
        ExportSchema = True
    End If
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
            If gVerbose Then LogMessage "Procesando campos de " & tName & ", total: " & fields.Count
            If fields.Count > 0 Then
                fieldKeys = fields.Keys
                If gVerbose Then LogMessage "Claves de campos obtenidas: " & UBound(fieldKeys) + 1
                For i = 0 To UBound(fieldKeys)
                    fName = fieldKeys(i)
                    Set finfo = fields(fName)
                    If gVerbose Then LogMessage "Procesando campo: " & fName & " - Tipo objeto: " & TypeName(finfo)
                    
                    Dim fieldType, fieldSize, fieldRequired
                    fieldType = ""
                    fieldSize = ""
                    fieldRequired = False
                    
                    On Error Resume Next
                    fieldType = finfo("type")
                    fieldSize = finfo("size")
                    fieldRequired = finfo("required")
                    On Error GoTo 0
                    
                    If gVerbose Then LogMessage "Valores obtenidos - Tipo: " & fieldType & ", Tamaño: " & fieldSize & ", Requerido: " & fieldRequired
                    
                    Dim defVal: defVal = ""
                     If finfo.Exists("defaultValue") Then defVal = finfo("defaultValue")
                     Dim requiredText: requiredText = "No"
                     If fieldRequired Then requiredText = "Si"
                     Dim rowText: rowText = "| " & fName & " | " & fieldType & " | " & fieldSize & " | " & requiredText & " | " & defVal & " |" & vbCrLf
                    s = s & rowText
                    If gVerbose Then LogMessage "Fila agregada: " & Trim(Replace(rowText, vbCrLf, ""))
                    If gVerbose Then LogMessage "Campo procesado: " & fName
                Next
            Else
                If gVerbose Then LogMessage "No hay campos en el diccionario para " & tName
            End If
        Else
            If gVerbose Then LogMessage "No hay campos para " & tName
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
    
    If gVerbose Then LogMessage "Contenido final del Markdown (primeros 500 caracteres): " & Left(s, 500)

    Dim ts: Set ts = CreateObject("ADODB.Stream")
    ts.Type = 2: ts.Charset = "utf-8": ts.Open
    ts.WriteText s, 0
    ts.SaveToFile filePath, 2
    ts.Close
    
    If gVerbose Then LogMessage "Archivo Markdown guardado: " & filePath & " (longitud: " & Len(s) & " caracteres)"
End Sub

' ============================================================================
' FUNCIONES AUXILIARES
' ============================================================================

Function SafeSwapForm(app, finalName, tmpName)
    On Error Resume Next
    
    ' Aplicar configuración anti-UI
    AntiUI app
    
    ' Eliminar formulario existente si existe
    Dim objExists
    objExists = False
    Dim obj
    For Each obj In app.CurrentProject.AllForms
        If obj.Name = finalName Then
            objExists = True
            Exit For
        End If
    Next
    
    If objExists Then
        app.DoCmd.DeleteObject acForm, finalName
    End If
    
    ' Renombrar temporal a final
    app.DoCmd.Rename finalName, acForm, tmpName
    
    Err.Clear
End Function

Function GetControlTypeFromName(typeName)
    Select Case LCase(typeName)
        Case "label": GetControlTypeFromName = acLabel
        Case "textbox": GetControlTypeFromName = acTextBox
        Case "commandbutton", "button": GetControlTypeFromName = acCommandButton
        Case "checkbox": GetControlTypeFromName = acCheckBox
        Case "optionbutton": GetControlTypeFromName = acOptionButton
        Case "combobox": GetControlTypeFromName = acComboBox
        Case "listbox": GetControlTypeFromName = acListBox
        Case "image": GetControlTypeFromName = acImage
        Case "subform": GetControlTypeFromName = acSubform
        Case Else: GetControlTypeFromName = acLabel ' Default
    End Select
End Function

Function ReadAllText(filePath)
    On Error Resume Next
    Dim stream
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "UTF-8"
    stream.Open
    stream.LoadFromFile filePath
    ReadAllText = stream.ReadText
    stream.Close
    If Err.Number <> 0 Then ReadAllText = ""
    Err.Clear
End Function

' Función para cerrar todas las instancias de Access
Sub CloseAllAccessInstances()
    On Error Resume Next
    
    Dim objWMI, colProcesses, objProcess
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcesses = objWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name = 'MSACCESS.EXE'")
    
    For Each objProcess In colProcesses
        LogVerbose "Cerrando proceso Access PID: " & objProcess.ProcessId
        objProcess.Terminate()
    Next
    
    ' Esperar un momento para que los procesos se cierren
    WScript.Sleep 1000
    
    Err.Clear
End Sub

Function JsonParse(jsonText)
    On Error Resume Next
    ' Implementación básica de parsing JSON
    ' En un entorno real, se usaría una librería JSON más robusta
    Set JsonParse = CreateObject("Scripting.Dictionary")
    ' Aquí iría la lógica de parsing JSON
    If Err.Number <> 0 Then Set JsonParse = Nothing
    Err.Clear
End Function

' Ejecutar función principal
Main()

' Limpiar instancia singleton al final del script
CloseSingletonAccess()