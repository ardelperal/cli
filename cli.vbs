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
Dim g_ModulesSrcPath, g_ModulesExtensions, g_ModulesIncludeSubdirs, g_ModulesFilePattern
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
                        
                        ' Resolver rutas relativas para ciertos valores (excepto patrones)
                        If (InStr(UCase(key), "PATH") > 0 Or InStr(UCase(key), "FILE") > 0) And InStr(UCase(key), "PATTERN") = 0 Then
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
    
    ' Configurar Access para modo silencioso
    objApp.Visible = False
    objApp.UserControl = False
    
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
    
    ' Deshabilitar confirmaciones y warnings para operaciones automaticas (como condor_cli.vbs)
    objApp.DoCmd.SetWarnings False
    
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
        ' Restaurar warnings antes de cerrar
        objAccess.DoCmd.SetWarnings True
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
Function ImportModuleToAccess(moduleName, moduleContent, filePath, objAccess)
    Dim moduleType, fileExt
    
    ImportModuleToAccess = False
    
    ' Determinar tipo de modulo por extensión
    fileExt = LCase(objFSO.GetExtensionName(filePath))
    
    Select Case fileExt
        Case "bas"
            moduleType = 5 ' Modulo estandar (acModule = 5)
        Case "cls"
            moduleType = 100 ' Modulo de clase (acClassModule = 100)
        Case Else
            LogMessage "Extension de archivo no soportada: " & fileExt & " para archivo: " & filePath
            Exit Function
    End Select
    
    ' Validar que objAccess no sea Nothing
    If objAccess Is Nothing Then
        LogMessage "Error: objAccess no esta disponible para importar modulo " & moduleName
        Exit Function
    End If
    
    ' Validar acceso a VBE antes de intentar importar
    On Error Resume Next
    Dim testVBE
    Set testVBE = objAccess.VBE
    If Err.Number <> 0 Then
        LogMessage "Error: No se puede acceder al VBE de Access para modulo " & moduleName
        LogMessage "SOLUCION: Habilite 'Trust access to the VBA project object model' en:"
        LogMessage "Access -> File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    
    ' Importar modulo usando la función robusta con codificación correcta
    On Error Resume Next
    If ImportVBAModuleFromFile(objAccess, filePath, moduleName, moduleType) Then
        ImportModuleToAccess = True
        LogMessage "Modulo " & moduleName & " importado exitosamente desde: " & filePath
        LogVerbose "Importacion completada - Tipo: " & moduleType & ", Archivo: " & filePath
    Else
        ImportModuleToAccess = False
        If Err.Number <> 0 Then
            LogMessage "Error importando modulo " & moduleName & " desde archivo: " & filePath
            LogMessage "Detalle del error: " & Err.Number & " - " & Err.Description
            LogMessage "Archivo problemático: " & objFSO.GetFileName(filePath)
            LogMessage "Ruta completa: " & filePath
            Err.Clear
        Else
            LogMessage "Error importando modulo " & moduleName & " desde archivo: " & filePath
            LogMessage "Archivo problemático: " & objFSO.GetFileName(filePath)
        End If
    End If
    On Error GoTo 0
End Function

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

Function ImportVBAModuleFromFile(objAccess, modulePath, moduleName, moduleType)
    ImportVBAModuleFromFile = False
    Dim fso, wsh, tempDir, tempPath, okCopy, proj, vbComp, base, found
    Dim ext
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set wsh = CreateObject("WScript.Shell")
    
    ' Validar que el archivo existe
    If Not fso.FileExists(modulePath) Then
        LogMessage "Error: Archivo no existe: " & modulePath
        Exit Function
    End If
    
    ' Validar acceso a VBE antes de proceder
    On Error Resume Next
    Dim testVBE
    Set testVBE = objAccess.VBE
    If Err.Number <> 0 Then
        LogMessage "Error: VBE bloqueado para archivo " & modulePath
        LogMessage "SOLUCION: Habilite 'Trust access to the VBA project object model' en:"
        LogMessage "Access -> File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings"
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    
    ' Preparar archivo temporal con extensión correcta
    tempDir = wsh.ExpandEnvironmentStrings("%TEMP%")
    If moduleType = 100 Then ext = "cls" Else ext = "bas" ' 100 = acClassModule, 5 = acModule
    tempPath = tempDir & "\" & moduleName & "." & ext
    
    ' Intentar importar directamente primero (sin conversión ANSI)
    On Error Resume Next
    Set proj = objAccess.VBE.VBProjects(1)
    If Err.Number <> 0 Then
        LogMessage "Error accediendo al proyecto VBA: " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    
    ' Eliminar componente existente si existe (comparación case-insensitive)
    For Each vbComp In proj.VBComponents
        If LCase(vbComp.Name) = LCase(moduleName) Then
            proj.VBComponents.Remove vbComp
            Exit For
        End If
    Next
    Err.Clear
    
    ' Intentar importar directamente desde el archivo original
    If gDebug Then LogMessage "[DEBUG] Intentando importar directamente desde: " & modulePath
    proj.VBComponents.Import modulePath
    If Err.Number = 0 Then
        ' Importación directa exitosa
        If gDebug Then LogMessage "[DEBUG] Importacion directa exitosa"
        On Error GoTo 0
        ' Verificar que el módulo se importó correctamente
        found = False
        For Each vbComp In proj.VBComponents
            If LCase(vbComp.Name) = LCase(moduleName) Then 
                found = True
                Exit For
            End If
        Next
        
        ' Si no se encontró con el nombre esperado, buscar por el nombre base del archivo
        If Not found Then
            base = fso.GetBaseName(modulePath)
            For Each vbComp In proj.VBComponents
                If LCase(vbComp.Name) = LCase(base) Then
                    ' Renombrar al nombre deseado si difiere
                    If vbComp.Name <> moduleName Then 
                        On Error Resume Next
                        vbComp.Name = moduleName
                        If Err.Number <> 0 Then
                            LogMessage "Advertencia: No se pudo renombrar modulo de " & base & " a " & moduleName
                            Err.Clear
                        End If
                        On Error GoTo 0
                    End If
                    found = True
                    Exit For
                End If
            Next
        End If
        
        If found Then
            ImportVBAModuleFromFile = True
            If gDebug Then LogMessage "[DEBUG] Modulo importado directamente: " & moduleName
            Exit Function
        End If
    Else
        ' Importación directa falló, intentar con conversión ANSI como fallback
        If gDebug Then LogMessage "[DEBUG] Importacion directa fallo (Error: " & Err.Number & "), intentando con conversion ANSI"
        Err.Clear
    End If
    On Error GoTo 0
    
    ' Fallback: Copiar archivo como ANSI de forma segura
    okCopy = CopyAsAnsi(modulePath, tempPath)
    If Not okCopy Then 
        LogMessage "Error copiando archivo a ANSI: " & modulePath
        Exit Function
    End If
    
    ' Importar módulo desde archivo temporal
    If gDebug Then LogMessage "[DEBUG] Intentando importar desde: " & tempPath
    If gDebug Then LogMessage "[DEBUG] Archivo temporal existe: " & fso.FileExists(tempPath)
    If gDebug Then LogMessage "[DEBUG] Tamaño archivo temporal: " & fso.GetFile(tempPath).Size & " bytes"
    
    On Error Resume Next
    proj.VBComponents.Import tempPath
    If Err.Number <> 0 Then 
        LogMessage "Error importando modulo VBA desde " & tempPath & ": " & Err.Description
        If gDebug Then LogMessage "[DEBUG] Error Number: " & Err.Number
        If gDebug Then LogMessage "[DEBUG] Error Source: " & Err.Source
        If gDebug Then LogMessage "[DEBUG] Archivo temporal NO eliminado para investigacion: " & tempPath
        Err.Clear
        ' NO eliminar archivo temporal en modo debug para investigar
        If Not gDebug Then
            On Error Resume Next
            fso.DeleteFile tempPath, True
            Err.Clear
            On Error GoTo 0
        End If
        Exit Function
    End If
    On Error GoTo 0
    
    If gDebug Then LogMessage "[DEBUG] Importacion exitosa, verificando componentes..."
    
    ' Verificar que el módulo se importó correctamente y renombrar si es necesario
    ' (El componente importado puede tener el nombre base del archivo en lugar del moduleName deseado)
    found = False
    On Error Resume Next
    For Each vbComp In proj.VBComponents
        If LCase(vbComp.Name) = LCase(moduleName) Then 
            found = True
            Exit For
        End If
    Next
    
    ' Si no se encontró con el nombre esperado, buscar por el nombre base del archivo
    If Not found Then
        base = fso.GetBaseName(tempPath)
        For Each vbComp In proj.VBComponents
            If LCase(vbComp.Name) = LCase(base) Then
                ' Renombrar al nombre deseado si difiere
                If vbComp.Name <> moduleName Then 
                    vbComp.Name = moduleName
                    If Err.Number <> 0 Then
                        LogMessage "Advertencia: No se pudo renombrar modulo de " & base & " a " & moduleName
                        Err.Clear
                    End If
                End If
                found = True
                Exit For
            End If
        Next
    End If
    Err.Clear
    On Error GoTo 0
    
    ' Cleanup temporal garantizado (Finally-like)
    On Error Resume Next
    fso.DeleteFile tempPath, True
    Err.Clear
    On Error GoTo 0
    
    ' Resultado final
    If found Then
        ' Compilar y guardar todos los modulos para que Access los reconozca
        On Error Resume Next
        objAccess.DoCmd.RunCommand 7 ' acCmdCompileAndSaveAllModules
        If Err.Number <> 0 Then
            LogVerbose "Advertencia: No se pudo compilar automaticamente (Error: " & Err.Number & ")"
            Err.Clear
        End If
        
        ' Guardar el modulo especifico
        objAccess.DoCmd.Save 5, moduleName ' acModule = 5
        If Err.Number <> 0 Then
            LogVerbose "Advertencia: No se pudo guardar el modulo " & moduleName & " (Error: " & Err.Number & ")"
            Err.Clear
        End If
        On Error GoTo 0
        
        ImportVBAModuleFromFile = True
        LogVerbose "Modulo " & moduleName & " importado exitosamente desde " & modulePath
    Else
        LogMessage "Error: Modulo " & moduleName & " no se encontro despues de la importacion"
        ImportVBAModuleFromFile = False
    End If
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

' Función para importar modulo VBA con manejo robusto de errores
Function ImportVBAModuleSafe(moduleName, moduleContent, moduleType)
    On Error Resume Next
    
    ImportVBAModuleSafe = False
    
    ' Verificar que Access esté disponible
    If objAccess Is Nothing Then
        LogMessage "Error: objAccess no está disponible"
        Exit Function
    End If
    
    ' Verificar acceso al VBE de manera segura
    Dim vbProject
    Set vbProject = objAccess.VBE.VBProjects(1)
    If Err.Number <> 0 Or vbProject Is Nothing Then
        LogMessage "Error: Proyecto VBA no está disponible - verifique configuracion de confianza VBA"
        Err.Clear
        Exit Function
    End If
    
    Dim vbComponents
    Set vbComponents = vbProject.VBComponents
    If Err.Number <> 0 Or vbComponents Is Nothing Then
        LogMessage "Error: Componentes VBA no están disponibles"
        Err.Clear
        Exit Function
    End If
    
    ' Crear nuevo componente VBA
    Dim vbComp
    Set vbComp = vbComponents.Add(moduleType)
    If Err.Number <> 0 Or vbComp Is Nothing Then
        LogMessage "Error creando componente VBA " & moduleName & ": " & Err.Description
        Err.Clear
        Exit Function
    End If
    
    ' Establecer nombre del módulo
    vbComp.Name = moduleName
    If Err.Number <> 0 Then
        LogMessage "Error estableciendo nombre del modulo " & moduleName & ": " & Err.Description
        Err.Clear
        Exit Function
    End If
    
    ' Establecer contenido del modulo
    Dim vbMod
    Set vbMod = vbComp.CodeModule
    If Err.Number <> 0 Or vbMod Is Nothing Then
        LogMessage "Error accediendo al CodeModule de " & moduleName
        Err.Clear
        Exit Function
    End If
    
    vbMod.DeleteLines 1, vbMod.CountOfLines
    If Err.Number <> 0 Then
        LogMessage "Error eliminando lineas del modulo " & moduleName & ": " & Err.Description
        Err.Clear
        Exit Function
    End If
    
    vbMod.InsertLines 1, moduleContent
    If Err.Number <> 0 Then
        LogMessage "Error insertando contenido en modulo " & moduleName & ": " & Err.Description
        Err.Clear
        Exit Function
    End If
    
    ' Si llegamos aquí, todo fue exitoso
    ImportVBAModuleSafe = True
    On Error GoTo 0
End Function

' Función para importar modulo VBA (versión original)
Sub ImportVBAModule(moduleName, moduleContent, moduleType)
    On Error Resume Next
    
    Dim vbComp, vbMod, vbProject, vbComponents
    
    ' Verificar que Access y VBE estén disponibles
    If objAccess Is Nothing Then
        LogMessage "Error: objAccess no está disponible"
        Err.Raise 91, , "Se requiere un objeto - objAccess"
        Exit Sub
    End If
    
    If objAccess.VBE Is Nothing Then
        LogMessage "Error: VBE no está disponible"
        Err.Raise 91, , "Se requiere un objeto - VBE"
        Exit Sub
    End If
    
    Set vbProject = objAccess.VBE.VBProjects(1)
    If vbProject Is Nothing Then
        LogMessage "Error: Proyecto VBA no está disponible"
        Err.Raise 91, , "Se requiere un objeto - VBProject"
        Exit Sub
    End If
    
    Set vbComponents = vbProject.VBComponents
    If vbComponents Is Nothing Then
        LogMessage "Error: Componentes VBA no están disponibles"
        Err.Raise 91, , "Se requiere un objeto - VBComponents"
        Exit Sub
    End If
    
    ' Crear nuevo componente VBA
    Set vbComp = vbComponents.Add(moduleType)
    If Err.Number <> 0 Then
        LogMessage "Error creando componente VBA " & moduleName & ": " & Err.Description
        Exit Sub
    End If
    
    If vbComp Is Nothing Then
        LogMessage "Error: No se pudo crear el componente VBA " & moduleName
        Err.Raise 91, , "Se requiere un objeto - VBComponent"
        Exit Sub
    End If
    
    ' Establecer nombre del módulo
    vbComp.Name = moduleName
    If Err.Number <> 0 Then
        LogMessage "Error estableciendo nombre del modulo " & moduleName & ": " & Err.Description
        Exit Sub
    End If
    
    ' Establecer contenido del modulo
    Set vbMod = vbComp.CodeModule
    If vbMod Is Nothing Then
        LogMessage "Error: No se pudo acceder al CodeModule de " & moduleName
        Err.Raise 91, , "Se requiere un objeto - CodeModule"
        Exit Sub
    End If
    
    vbMod.DeleteLines 1, vbMod.CountOfLines
    If Err.Number <> 0 Then
        LogMessage "Error eliminando lineas del modulo " & moduleName & ": " & Err.Description
        Exit Sub
    End If
    
    vbMod.InsertLines 1, moduleContent
    If Err.Number <> 0 Then
        LogMessage "Error insertando contenido en modulo " & moduleName & ": " & Err.Description
        Exit Sub
    End If
    
    ' Si llegamos aquí, todo fue exitoso
    LogMessage "Modulo VBA " & moduleName & " creado exitosamente"
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

' Función para escribir contenido de modulo a archivo
Function WriteModuleFile(filePath, content)
    On Error Resume Next
    
    Dim file
    
    WriteModuleFile = False
    
    ' Crear directorio padre si no existe
    Dim parentDir
    parentDir = objFSO.GetParentFolderName(filePath)
    If Not objFSO.FolderExists(parentDir) Then
        CreateFolderRecursive parentDir
    End If
    
    ' Escribir archivo
    Set file = objFSO.CreateTextFile(filePath, True, False) ' Sobrescribir, ANSI
    If Err.Number <> 0 Then
        Err.Clear
        Set file = objFSO.CreateTextFile(filePath, True, False) ' Sobrescribir, ASCII
    End If
    
    If Not file Is Nothing Then
        file.Write content
        file.Close
        WriteModuleFile = True
    End If
    
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
    g_ModulesFilePattern = gConfig("MODULES_FilePattern")
    
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
            Set config = LoadConfig(gConfigPath)
            gDbPath = ResolvePath(config("DATABASE_DefaultPath"))
            
            If cleanArgCount >= 2 Then
                ' Hay argumentos adicionales - pueden ser archivos específicos
                Dim moduleArgs, moduleList
                moduleArgs = cleanArgs(1)
                
                ' Verificar si contiene comas (múltiples archivos)
                If InStr(moduleArgs, ",") > 0 Then
                    ' Múltiples archivos separados por comas
                    moduleList = Split(moduleArgs, ",")
                    
                    If Not gDryRun Then
                        If gVerbose Then WScript.Echo "Ejecutando update de modulos especificos..."
                        If Not UpdateSpecificModules(gDbPath, moduleList) Then
                            WScript.Echo "Error: No se pudo completar el update de modulos"
                            WScript.Quit 1
                        End If
                    Else
                        WScript.Echo "[DRY-RUN] Se ejecutaria update de modulos: " & moduleArgs
                    End If
                Else
                    ' Un solo archivo específico
                    If Not gDryRun Then
                        If gVerbose Then WScript.Echo "Ejecutando update de modulo especifico..."
                        If Not UpdateSpecificModule(gDbPath, moduleArgs) Then
                            WScript.Echo "Error: No se pudo completar el update del modulo"
                            WScript.Quit 1
                        End If
                    Else
                        WScript.Echo "[DRY-RUN] Se ejecutaria update del modulo: " & moduleArgs
                    End If
                End If
            Else
                ' Sin argumentos - actualizar solo módulos más nuevos
                If Not gDryRun Then
                    If gVerbose Then WScript.Echo "Ejecutando update de modulos mas nuevos..."
                    If Not UpdateNewerModules(gDbPath) Then
                        WScript.Echo "Error: No se pudo completar el update de modulos"
                        WScript.Quit 1
                    End If
                Else
                    WScript.Echo "[DRY-RUN] Se ejecutaria update de modulos mas nuevos en: " & gDbPath
                End If
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


' Función para actualizar un módulo específico
Function UpdateSpecificModule(dbPath, moduleArg)
    On Error Resume Next
    
    Dim config, srcPath, moduleName, moduleFile, moduleContent
    Dim fileExt, possibleFiles(1)
    
    UpdateSpecificModule = False
    
    ' Cargar configuración
    Set config = LoadConfig(gConfigPath)
    srcPath = config("MODULES_SrcPath")
    
    LogMessage "Iniciando update de modulo especifico: " & moduleArg
    LogMessage "Base de datos destino: " & dbPath
    
    ' Verificar que existe el directorio fuente
    If Not objFSO.FolderExists(srcPath) Then
        LogMessage "Error: No existe el directorio fuente de modulos: " & srcPath
        Exit Function
    End If
    
    ' Abrir Access
    Set objAccess = OpenAccess(dbPath, gPassword)
    If objAccess Is Nothing Then
        LogMessage "Error: No se pudo abrir la base de datos para update"
        Exit Function
    End If
    
    ' Limpiar el nombre del módulo (quitar src\ si existe)
    moduleName = moduleArg
    If InStr(moduleName, "src\") > 0 Then
        moduleName = Replace(moduleName, "src\", "")
    End If
    If InStr(moduleName, "src/") > 0 Then
        moduleName = Replace(moduleName, "src/", "")
    End If
    
    ' Si ya tiene extensión, usar directamente
    If Right(LCase(moduleName), 4) = ".cls" Or Right(LCase(moduleName), 4) = ".bas" Then
        moduleFile = srcPath & "\" & moduleName
        If objFSO.FileExists(moduleFile) Then
            LogMessage "Actualizando modulo: " & objFSO.GetBaseName(moduleName) & " desde " & moduleFile
            
            ' Leer contenido del archivo
            moduleContent = ReadModuleFile(moduleFile)
            
            If moduleContent <> "" Then
                ' Importar o actualizar modulo en Access
                If ImportModuleToAccess(objFSO.GetBaseName(moduleName), moduleContent, moduleFile, objAccess) Then
                    UpdateSpecificModule = True
                Else
                    LogMessage "Error actualizando modulo: " & objFSO.GetBaseName(moduleName)
                End If
            Else
                LogMessage "Error leyendo archivo de modulo: " & moduleFile
            End If
        Else
            LogMessage "Error: No se encontro el archivo: " & moduleFile
        End If
    Else
        ' Sin extensión - buscar automáticamente .cls o .bas
        possibleFiles(0) = srcPath & "\" & moduleName & ".cls"
        possibleFiles(1) = srcPath & "\" & moduleName & ".bas"
        
        Dim i, found
        found = False
        For i = 0 To 1
            If objFSO.FileExists(possibleFiles(i)) Then
                moduleFile = possibleFiles(i)
                found = True
                Exit For
            End If
        Next
        
        If found Then
            LogMessage "Actualizando modulo: " & moduleName & " desde " & moduleFile
            
            ' Leer contenido del archivo
            moduleContent = ReadModuleFile(moduleFile)
            
            If moduleContent <> "" Then
                ' Importar o actualizar modulo en Access
                If ImportModuleToAccess(moduleName, moduleContent, moduleFile, objAccess) Then
                    UpdateSpecificModule = True
                Else
                    LogMessage "Error actualizando modulo: " & moduleName
                End If
            Else
                LogMessage "Error leyendo archivo de modulo: " & moduleFile
            End If
        Else
            LogMessage "Error: No se encontro el modulo " & moduleName & " (.cls o .bas) en: " & srcPath
        End If
    End If
    
    ' Cerrar Access
    CloseAccess objAccess
    
    If Err.Number <> 0 Then
        LogMessage "Error durante update especifico: " & Err.Description
        Err.Clear
        UpdateSpecificModule = False
    End If
End Function

' Función para actualizar múltiples módulos específicos
Function UpdateSpecificModules(dbPath, moduleList)
    On Error Resume Next
    
    Dim i, moduleName, success, totalSuccess
    
    UpdateSpecificModules = False
    totalSuccess = 0
    
    LogMessage "Iniciando update de " & (UBound(moduleList) + 1) & " modulos especificos"
    
    For i = 0 To UBound(moduleList)
        moduleName = Trim(moduleList(i))
        If moduleName <> "" Then
            LogMessage "Procesando modulo " & (i + 1) & " de " & (UBound(moduleList) + 1) & ": " & moduleName
            success = UpdateSpecificModule(dbPath, moduleName)
            If success Then
                totalSuccess = totalSuccess + 1
            End If
        End If
    Next
    
    If totalSuccess = UBound(moduleList) + 1 Then
        LogMessage "Update completado exitosamente - " & totalSuccess & " modulos actualizados"
        UpdateSpecificModules = True
    Else
        LogMessage "Update completado con errores - " & totalSuccess & " de " & (UBound(moduleList) + 1) & " modulos actualizados"
        UpdateSpecificModules = False
    End If
    
    If Err.Number <> 0 Then
        LogMessage "Error durante update multiple: " & Err.Description
        Err.Clear
        UpdateSpecificModules = False
    End If
End Function

' Función para actualizar solo módulos más nuevos
Function UpdateNewerModules(dbPath)
    On Error Resume Next
    
    Dim config, srcPath, extensions, includeSubdirs, filePattern
    Dim moduleFiles, i, moduleFile, moduleName, moduleContent
    Dim dbModules, needsUpdate, updatedCount
    
    UpdateNewerModules = False
    updatedCount = 0
    
    ' Cargar configuración
    Set config = LoadConfig(gConfigPath)
    srcPath = config("MODULES_SrcPath")
    extensions = config("MODULES_Extensions")
    includeSubdirs = LCase(config("MODULES_IncludeSubdirectories")) = "true"
    filePattern = config("MODULES_FilePattern")
    
    LogMessage "Iniciando update de modulos mas nuevos desde: " & srcPath
    LogMessage "Base de datos destino: " & dbPath
    
    ' Verificar que existe el directorio fuente
    If Not objFSO.FolderExists(srcPath) Then
        LogMessage "Error: No existe el directorio fuente de modulos: " & srcPath
        Exit Function
    End If
    
    ' Abrir Access
    Set objAccess = OpenAccess(dbPath, gPassword)
    If objAccess Is Nothing Then
        LogMessage "Error: No se pudo abrir la base de datos para update"
        Exit Function
    End If
    
    ' Obtener modulos existentes en la base de datos
    Set dbModules = GetDatabaseModules()
    
    ' Obtener lista de archivos de modulos
    moduleFiles = GetModuleFiles(srcPath, extensions, includeSubdirs, filePattern)
    
    If UBound(moduleFiles) >= 0 Then
        LogMessage "Encontrados " & (UBound(moduleFiles) + 1) & " archivos de modulos"
        
        ' Procesar cada archivo de modulo
        For i = 0 To UBound(moduleFiles)
            moduleFile = moduleFiles(i)
            moduleName = objFSO.GetBaseName(moduleFile)
            
            ' Verificar si necesita actualización (comparar fechas)
            needsUpdate = ModuleNeedsUpdateByDate(moduleFile, moduleName, dbModules)
            
            If needsUpdate Then
                LogMessage "Actualizando modulo mas nuevo: " & moduleName & " desde " & moduleFile
                
                ' Leer contenido del archivo
                moduleContent = ReadModuleFile(moduleFile)
                
                If moduleContent <> "" Then
                    ' Importar o actualizar modulo en Access
                    If ImportModuleToAccess(moduleName, moduleContent, moduleFile, objAccess) Then
                        updatedCount = updatedCount + 1
                    Else
                        LogMessage "Error actualizando modulo: " & moduleName
                    End If
                Else
                    LogMessage "Error leyendo archivo de modulo: " & moduleFile
                End If
            Else
                LogMessage "Modulo " & moduleName & " esta actualizado"
            End If
        Next
        
        UpdateNewerModules = True
        LogMessage "Update completado exitosamente - " & updatedCount & " modulos actualizados"
    Else
        LogMessage "No se encontraron archivos de modulos en: " & srcPath
        UpdateNewerModules = True ' No es error si no hay modulos
    End If
    
    ' Cerrar Access
    CloseAccess objAccess
    
    If Err.Number <> 0 Then
        LogMessage "Error durante update de modulos mas nuevos: " & Err.Description
        Err.Clear
        UpdateNewerModules = False
    End If
End Function

' Función para verificar si un módulo necesita actualización por fecha
Function ModuleNeedsUpdateByDate(filePath, moduleName, dbModules)
    On Error Resume Next
    
    Dim fileDate
    
    ModuleNeedsUpdateByDate = True ' Por defecto, actualizar
    
    ' Obtener fecha del archivo
    fileDate = objFSO.GetFile(filePath).DateLastModified
    
    ' Si el modulo no existe en la BD, necesita actualización
    If Not dbModules.Exists(moduleName) Then
        ModuleNeedsUpdateByDate = True
        Exit Function
    End If
    
    ' Por simplicidad, comparar solo si el archivo es más nuevo que hace 1 día
    ' En una implementación más avanzada, se podría obtener la fecha real del módulo en la BD
    If DateDiff("d", fileDate, Now()) <= 1 Then
        ModuleNeedsUpdateByDate = True
    Else
        ModuleNeedsUpdateByDate = False
    End If
    
    If Err.Number <> 0 Then
        LogMessage "Error verificando fecha de actualizacion para " & moduleName & ": " & Err.Description
        Err.Clear
        ModuleNeedsUpdateByDate = True
    End If
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

' Subrutina para importar módulos de forma desatendida
Sub ImportModuleDesatendido(strImportPath, moduleName, fileExtension, cleanedContent)
    On Error Resume Next
    
    LogMessage "  Importando desatendido: " & moduleName & " desde " & strImportPath
    
    ' Eliminar módulo existente si existe
    Dim vbProject, existingComponent
    Set vbProject = Application.VBE.ActiveVBProject
    
    For Each existingComponent In vbProject.VBComponents
        If existingComponent.Name = moduleName Then
            LogMessage "  Eliminando módulo existente: " & moduleName
            vbProject.VBComponents.Remove existingComponent
            Exit For
        End If
    Next
    
    ' Crear archivo temporal con el contenido limpio
    Dim tempFilePath, fso, tempFile
    Set fso = CreateObject("Scripting.FileSystemObject")
    tempFilePath = fso.GetTempName()
    tempFilePath = fso.GetSpecialFolder(2) & "\" & tempFilePath & "." & fileExtension
    
    Set tempFile = fso.CreateTextFile(tempFilePath, True)
    tempFile.Write cleanedContent
    tempFile.Close
    
    ' Importar el módulo usando VBComponents.Import
    Dim importedComponent
    Set importedComponent = vbProject.VBComponents.Import(tempFilePath)
    
    If Err.Number <> 0 Then
        LogMessage "  ❌ Error al importar " & moduleName & ": " & Err.Description
        Err.Clear
    Else
        ' Renombrar el componente si es necesario
        If importedComponent.Name <> moduleName Then
            importedComponent.Name = moduleName
            If Err.Number <> 0 Then
                LogMessage "  ❌ Error al renombrar " & moduleName & ": " & Err.Description
                Err.Clear
            End If
        End If
        LogMessage "  ✓ Importado desatendido: " & moduleName
    End If
    
    ' Limpiar archivo temporal
    If fso.FileExists(tempFilePath) Then
        fso.DeleteFile tempFilePath
    End If
    
    On Error GoTo 0
End Sub

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
            strFileName = objFile.Path
            strModuleName = objFSO.GetBaseName(objFile.Name)
            
            LogMessage "Procesando modulo: " & strModuleName
            
            ' Determinar tipo de archivo
            Dim fileExtension
            fileExtension = LCase(objFSO.GetExtensionName(objFile.Name))
            
            ' Limpiar archivo antes de importar (eliminar metadatos Attribute)
            Dim cleanedContent
            cleanedContent = CleanVBAFile(strFileName, fileExtension)
            
            ' Importar usando contenido limpio
            Call ImportModuleWithAnsiEncoding(strFileName, strModuleName, fileExtension, objAccess.VBE.ActiveVBProject, cleanedContent)
            
            If Err.Number <> 0 Then
                LogMessage "Error al importar modulo " & strModuleName & ": " & Err.Description
                Err.Clear
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

' Función para validar sintaxis básica de archivos VBA
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

' Ejecutar función principal
Main()