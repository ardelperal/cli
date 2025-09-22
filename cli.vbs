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
Dim gVerbose, gQuiet, gTestMode, gDryRun
Dim gDbPath, gPassword, gOutputPath, gConfigPath, gScriptPath, gScriptDir, gTimeout
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
                        
                        ' Resolver rutas relativas para ciertos valores
                        If InStr(UCase(key), "PATH") > 0 Or InStr(UCase(key), "FILE") > 0 Then
                            If value <> "" And fso.GetAbsolutePathName(value) <> value Then
                                value = gScriptDir & "\" & value
                            End If
                        End If
                        
                        If config.Exists(key) Then
                            config(key) = value
                        Else
                            config.Add key, value
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
    WScript.Echo "  extract-all <db_path> [output_path]"
    WScript.Echo "    Extrae toda la informacion de la base de datos"
    WScript.Echo ""
    WScript.Echo "  extract-tables <db_path> [output_path]"
    WScript.Echo "    Extrae informacion de tablas, campos y relaciones"
    WScript.Echo ""
    WScript.Echo "  extract-forms <db_path> [output_path]"
    WScript.Echo "    Extrae informacion de formularios y controles"
    WScript.Echo ""
    WScript.Echo "  extract-modules <db_path>"
    WScript.Echo "    Extrae modulos VBA hacia archivos fuente"
    WScript.Echo ""
    WScript.Echo "  list-objects <db_path>"
    WScript.Echo "    Lista todos los objetos de la base de datos"
    WScript.Echo ""
    WScript.Echo "  rebuild <db_path>"
    WScript.Echo "    Reconstruir modulos VBA desde src"
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
    WScript.Echo "  cscript cli.vbs extract-all ""C:\mi_base.accdb"""
    WScript.Echo "  cscript cli.vbs extract-tables ""C:\mi_base.accdb"" ""C:\output"""
    WScript.Echo "  cscript cli.vbs list-objects ""C:\mi_base.accdb"" --verbose"
    WScript.Echo "  cscript cli.vbs rebuild ""C:\mi_base.accdb"""
    WScript.Echo "  cscript cli.vbs update ""C:\mi_base.accdb"" /verbose"
    WScript.Echo "  cscript cli.vbs extract-all ""C:\mi_base.accdb"""
End Sub

' ============================================================================
' SECCIÓN 5: FUNCIONES DE ACCESS
' ============================================================================

' Función para abrir Access de forma segura
' Función canónica para abrir Access (basada en condor_cli.vbs)
Function OpenAccess(dbPath, password)
    Dim objAccess
    
    If gVerbose Then
        WScript.Echo "[VERBOSE] Abriendo Access: " & dbPath
        If password <> "" Then
            WScript.Echo "[VERBOSE] Con password: (oculta)"
        End If
    End If
    
    On Error Resume Next
    Set objAccess = CreateObject("Access.Application")
    
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: No se pudo crear instancia de Access: " & Err.Description
        Set OpenAccess = Nothing
        Exit Function
    End If
    
    ' Configurar Access para modo silencioso
    objAccess.Visible = False
    objAccess.UserControl = False
    
    ' Abrir la base de datos
    If password <> "" Then
        objAccess.OpenCurrentDatabase dbPath, False, password
    Else
        objAccess.OpenCurrentDatabase dbPath, False
    End If
    
    If Err.Number <> 0 Then
        WScript.Echo "ERROR: No se pudo abrir la base de datos: " & Err.Description
        objAccess.Quit
        Set objAccess = Nothing
        Set OpenAccess = Nothing
        Exit Function
    End If
    
    On Error GoTo 0
    Set OpenAccess = objAccess
    
    If gVerbose Then
        ' Verificar si VBE está disponible
        Dim vbeAvailable
        vbeAvailable = False
        On Error Resume Next
        If Not objAccess.VBE Is Nothing Then
            vbeAvailable = True
        End If
        On Error GoTo 0
        
        If vbeAvailable Then
            WScript.Echo "[VERBOSE] Acceso VBE disponible"
        Else
            WScript.Echo "[VERBOSE] Acceso VBE NO disponible"
        End If
        
        WScript.Echo "[VERBOSE] Access abierto exitosamente"
    End If
End Function

' Función canónica para cerrar Access de forma segura (basada en condor_cli.vbs)
Sub CloseAccess(objApp)
    If Not objApp Is Nothing Then
        LogVerbose "Cerrando Access..."
        
        On Error Resume Next
        objApp.CloseCurrentDatabase
        objApp.Quit
        Set objApp = Nothing
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

' Funciones para extraer información de tablas
Function ExtractTablesInfo(db)
    Dim tablesDict, tableDict, tbl, fld, idx, rel
    Set tablesDict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    
    ' Recorrer todas las tablas
    For Each tbl In db.TableDefs
        ' Filtrar tablas del sistema
        If Left(tbl.Name, 4) <> "MSys" And Left(tbl.Name, 1) <> "~" Then
            Set tableDict = CreateObject("Scripting.Dictionary")
            
            ' Información básica de la tabla
            tableDict("name") = tbl.Name
            tableDict("recordCount") = tbl.RecordCount
            tableDict("dateCreated") = FormatDateTime(tbl.DateCreated, vbGeneralDate)
            tableDict("lastUpdated") = FormatDateTime(tbl.LastUpdated, vbGeneralDate)
            
            ' Campos de la tabla
            Dim fieldsArray()
            ReDim fieldsArray(tbl.Fields.Count - 1)
            Dim i: i = 0
            
            For Each fld In tbl.Fields
                Dim fieldDict
                Set fieldDict = CreateObject("Scripting.Dictionary")
                fieldDict("name") = fld.Name
                fieldDict("type") = GetFieldTypeName(fld.Type)
                fieldDict("size") = fld.Size
                fieldDict("required") = fld.Required
                fieldDict("allowZeroLength") = fld.AllowZeroLength
                If fld.DefaultValue <> "" Then
                    fieldDict("defaultValue") = fld.DefaultValue
                End If
                Set fieldsArray(i) = fieldDict
                i = i + 1
            Next
            
            tableDict("fields") = fieldsArray
            
            ' Índices de la tabla
            Dim indexesArray()
            If tbl.Indexes.Count > 0 Then
                ReDim indexesArray(tbl.Indexes.Count - 1)
                i = 0
                For Each idx In tbl.Indexes
                    Dim indexDict
                    Set indexDict = CreateObject("Scripting.Dictionary")
                    indexDict("name") = idx.Name
                    indexDict("primary") = idx.Primary
                    indexDict("unique") = idx.Unique
                    indexDict("required") = idx.Required
                    
                    ' Campos del índice
                    Dim indexFieldsArray()
                    ReDim indexFieldsArray(idx.Fields.Count - 1)
                    Dim j: j = 0
                    For Each fld In idx.Fields
                        indexFieldsArray(j) = fld.Name
                        j = j + 1
                    Next
                    indexDict("fields") = indexFieldsArray
                    
                    Set indexesArray(i) = indexDict
                    i = i + 1
                Next
            End If
            tableDict("indexes") = indexesArray
            
            Set tablesDict(tbl.Name) = tableDict
        End If
    Next
    
    Set ExtractTablesInfo = tablesDict
End Function

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

' Función principal de extracción completa
Function ExtractAll(dbPath, outputPath)
    LogMessage "Iniciando extraccion completa de: " & objFSO.GetFileName(dbPath)
    
    ' Convertir rutas relativas a absolutas basadas en el directorio del script
    If Not objFSO.IsAbsolutePathName(dbPath) Then
        dbPath = gScriptDir & "\" & dbPath
    End If
    
    If Not objFSO.IsAbsolutePathName(outputPath) Then
        outputPath = gScriptDir & "\" & outputPath
    End If
    
    Set objAccess = OpenAccess(dbPath, gPassword)
    If objAccess Is Nothing Then
        ExtractAll = False
        Exit Function
    End If
    
    ' Crear directorio de salida si no existe
    If Not objFSO.FolderExists(outputPath) Then
        CreateFolderRecursive outputPath
    End If
    
    ' Extraer información
    ExtractTables outputPath
    ExtractForms outputPath
    ExtractQueries outputPath
    ExtractReports outputPath
    ExtractModules outputPath
    ExtractRelationships outputPath
    
    CloseAccess objAccess
    
    LogMessage "Extraccion completa finalizada"
    ExtractAll = True
End Function

' Función para extraer información de tablas
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

' Funciones para extraer informacion de formularios
Function ExtractFormsInfo(db)
    Dim formsDict, formDict, doc, frm
    Set formsDict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    
    ' Recorrer todos los formularios
    For Each doc In db.Containers("Forms").Documents
        Set formDict = CreateObject("Scripting.Dictionary")
        
        ' Información básica del formulario
        formDict("name") = doc.Name
        formDict("dateCreated") = FormatDateTime(doc.DateCreated, vbGeneralDate)
        formDict("lastUpdated") = FormatDateTime(doc.LastUpdated, vbGeneralDate)
        formDict("owner") = doc.Owner
        
        ' Intentar abrir el formulario en modo diseño para obtener más información
        Dim app
        Set app = CreateObject("Access.Application")
        app.Visible = False
        app.OpenCurrentDatabase db.Name
        
        On Error Resume Next
        app.DoCmd.OpenForm doc.Name, acDesign, , , , acHidden
        
        If Err.Number = 0 Then
            Set frm = app.Forms(doc.Name)
            
            ' Propiedades del formulario
            formDict("recordSource") = frm.RecordSource
            formDict("caption") = frm.Caption
            formDict("defaultView") = GetViewTypeName(frm.DefaultView)
            formDict("allowEdits") = frm.AllowEdits
            formDict("allowDeletions") = frm.AllowDeletions
            formDict("allowAdditions") = frm.AllowAdditions
            formDict("dataEntry") = frm.DataEntry
            formDict("modal") = frm.Modal
            formDict("popUp") = frm.PopUp
            
            ' Extraer controles del formulario
            formDict("controls") = ExtractFormControlsInternal(frm)
            
            app.DoCmd.Close acForm, doc.Name, acSaveNo
        Else
            ' Si no se puede abrir, solo información básica
            formDict("recordSource") = ""
            formDict("caption") = ""
            formDict("controls") = Array()
        End If
        
        app.Quit
        Set app = Nothing
        On Error GoTo 0
        
        Set formsDict(doc.Name) = formDict
    Next
    
    Set ExtractFormsInfo = formsDict
End Function

Function ExtractFormControls(db, formName)
    Dim controlsDict, app, frm
    Set controlsDict = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    
    ' Abrir Access y el formulario
    Set app = CreateObject("Access.Application")
    app.Visible = False
    app.OpenCurrentDatabase db.Name
    
    app.DoCmd.OpenForm formName, acDesign, , , , acHidden
    
    If Err.Number <> 0 Then
        LogMessage "Error: Formulario '" & formName & "' no encontrado o no se puede abrir", "ERROR"
        app.Quit
        Set app = Nothing
        Set ExtractFormControls = controlsDict
        Exit Function
    End If
    
    Set frm = app.Forms(formName)
    Set controlsDict = ExtractFormControlsInternal(frm)
    
    app.DoCmd.Close acForm, formName, acSaveNo
    app.Quit
    Set app = Nothing
    
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

' Función para extraer información de formularios
Sub ExtractForms(outputPath)
    LogMessage "Extrayendo informacion de formularios..."
    
    Dim formInfo, frm
    Set formInfo = CreateObject("Scripting.Dictionary")
    
    ' Iterar sobre todos los formularios
    For Each frm In objAccess.CurrentProject.AllForms
        LogVerbose "Procesando formulario: " & frm.Name
        
        Dim formData
        Set formData = CreateObject("Scripting.Dictionary")
        
        formData.Add "Name", frm.Name
        formData.Add "IsLoaded", frm.IsLoaded
        
        ' Abrir formulario en modo diseño para acceder a controles
        On Error Resume Next
        objAccess.DoCmd.OpenForm frm.Name, 0 ' Modo diseño
        
        If Err.Number = 0 Then
            Dim controls, ctl
            Set controls = CreateObject("Scripting.Dictionary")
            
            For Each ctl In objAccess.Forms(frm.Name).Controls
                Dim controlData
                Set controlData = CreateObject("Scripting.Dictionary")
                
                controlData.Add "ControlType", ctl.ControlType
                controlData.Add "Left", ctl.Left
                controlData.Add "Top", ctl.Top
                controlData.Add "Width", ctl.Width
                controlData.Add "Height", ctl.Height
                
                controls.Add ctl.Name, controlData
            Next
            
            formData.Add "Controls", controls
            objAccess.DoCmd.Close 2, frm.Name ' Cerrar formulario
        End If
        
        On Error GoTo 0
        formInfo.Add frm.Name, formData
    Next
    
    ' Guardar información en archivo JSON
    SaveToJSON formInfo, outputPath & "\forms.json"
    LogMessage "Información de formularios guardada en: forms.json"
End Sub

' Función para extraer consultas
Sub ExtractQueries(outputPath)
    LogMessage "Extrayendo información de consultas..."
    
    Dim queryInfo, qry
    Set queryInfo = CreateObject("Scripting.Dictionary")
    
    For Each qry In objAccess.CurrentDb.QueryDefs
        LogVerbose "Procesando consulta: " & qry.Name
        
        Dim queryData
        Set queryData = CreateObject("Scripting.Dictionary")
        
        queryData.Add "Name", qry.Name
        queryData.Add "SQL", qry.SQL
        queryData.Add "Type", qry.Type
        queryData.Add "DateCreated", qry.DateCreated
        queryData.Add "LastUpdated", qry.LastUpdated
        
        queryInfo.Add qry.Name, queryData
    Next
    
    SaveToJSON queryInfo, outputPath & "\queries.json"
    LogMessage "Información de consultas guardada en: queries.json"
End Sub

' Función para extraer reportes
Sub ExtractReports(outputPath)
    LogMessage "Extrayendo información de reportes..."
    
    Dim reportInfo, rpt
    Set reportInfo = CreateObject("Scripting.Dictionary")
    
    For Each rpt In objAccess.CurrentProject.AllReports
        LogVerbose "Procesando reporte: " & rpt.Name
        
        Dim reportData
        Set reportData = CreateObject("Scripting.Dictionary")
        
        reportData.Add "Name", rpt.Name
        reportData.Add "IsLoaded", rpt.IsLoaded
        
        reportInfo.Add rpt.Name, reportData
    Next
    
    SaveToJSON reportInfo, outputPath & "\reports.json"
    LogMessage "Información de reportes guardada en: reports.json"
End Sub

' Función para extraer modulos
Sub ExtractModules(outputPath)
    LogMessage "Extrayendo informacion de modulos..."
    
    Dim moduleInfo, mdl
    Set moduleInfo = CreateObject("Scripting.Dictionary")
    
    For Each mdl In objAccess.CurrentProject.AllModules
        LogVerbose "Procesando modulo: " & mdl.Name
        
        Dim moduleData
        Set moduleData = CreateObject("Scripting.Dictionary")
        
        moduleData.Add "Name", mdl.Name
        moduleData.Add "IsLoaded", mdl.IsLoaded
        
        moduleInfo.Add mdl.Name, moduleData
    Next
    
    SaveToJSON moduleInfo, outputPath & "\modules.json"
    LogMessage "Informacion de modulos guardada en: modules.json"
End Sub

' Función para extraer relaciones
Sub ExtractRelationships(outputPath)
    LogMessage "Extrayendo información de relaciones..."
    
    Dim relationInfo, rel
    Set relationInfo = CreateObject("Scripting.Dictionary")
    
    For Each rel In objAccess.CurrentDb.Relations
        LogVerbose "Procesando relación: " & rel.Name
        
        Dim relationData
        Set relationData = CreateObject("Scripting.Dictionary")
        
        relationData.Add "Name", rel.Name
        relationData.Add "Table", rel.Table
        relationData.Add "ForeignTable", rel.ForeignTable
        relationData.Add "Attributes", rel.Attributes
        
        relationInfo.Add rel.Name, relationData
    Next
    
    SaveToJSON relationInfo, outputPath & "\relationships.json"
    LogMessage "Información de relaciones guardada en: relationships.json"
End Sub

' ============================================================================
' SECCIÓN 8: FUNCIONES UTILITARIAS
' ============================================================================

' Funciones utilitarias
Function GetFieldTypeName(fieldType)
    Select Case fieldType
        Case dbBoolean: GetFieldTypeName = "Boolean"
        Case dbByte: GetFieldTypeName = "Byte"
        Case dbInteger: GetFieldTypeName = "Integer"
        Case dbLong: GetFieldTypeName = "Long"
        Case dbCurrency: GetFieldTypeName = "Currency"
        Case dbSingle: GetFieldTypeName = "Single"
        Case dbDouble: GetFieldTypeName = "Double"
        Case dbDate: GetFieldTypeName = "Date/Time"
        Case dbBinary: GetFieldTypeName = "Binary"
        Case dbText: GetFieldTypeName = "Text"
        Case dbLongBinary: GetFieldTypeName = "OLE Object"
        Case dbMemo: GetFieldTypeName = "Memo"
        Case dbGUID: GetFieldTypeName = "Replication ID"
        Case dbBigInt: GetFieldTypeName = "Big Integer"
        Case dbVarBinary: GetFieldTypeName = "VarBinary"
        Case dbChar: GetFieldTypeName = "Char"
        Case dbNumeric: GetFieldTypeName = "Numeric"
        Case dbDecimal: GetFieldTypeName = "Decimal"
        Case dbFloat: GetFieldTypeName = "Float"
        Case dbTime: GetFieldTypeName = "Time"
        Case dbTimeStamp: GetFieldTypeName = "TimeStamp"
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
    
    Set outputFile = objFSO.CreateTextFile(filePath, True)
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
Function ModuleNeedsUpdate(filePath, moduleName, dbModules)
    On Error Resume Next
    
    Dim fileDate, dbDate
    
    ModuleNeedsUpdate = True ' Por defecto, actualizar
    
    ' Obtener fecha del archivo
    fileDate = objFSO.GetFile(filePath).DateLastModified
    
    ' Si el modulo no existe en la BD, necesita actualización
    If Not dbModules.Exists(moduleName) Then
        ModuleNeedsUpdate = True
        Exit Function
    End If
    
    ' Por simplicidad, siempre actualizar
    ' En una implementación más avanzada, se podría comparar fechas o contenido
    ModuleNeedsUpdate = True
    
    If Err.Number <> 0 Then
        LogMessage "Error verificando actualización para " & moduleName & ": " & Err.Description
        Err.Clear
        ModuleNeedsUpdate = True
    End If
End Function
' ============================================================================

' Función para reconstruir modulos VBA desde archivos fuente
Function RebuildModules(dbPath)
    On Error Resume Next
    
    Dim config, srcPath, extensions, includeSubdirs, filePattern
    Dim moduleFiles, i, moduleFile, moduleName, moduleContent
    
    RebuildModules = False
    
    ' Cargar configuración
    Set config = LoadConfig(gConfigPath)
    srcPath = config("MODULES_SrcPath")
    extensions = config("MODULES_Extensions")
    includeSubdirs = LCase(config("MODULES_IncludeSubdirectories")) = "true"
    filePattern = config("MODULES_FilePattern")
    
    LogMessage "Iniciando rebuild de modulos VBA desde: " & srcPath
    LogMessage "Base de datos destino: " & dbPath
    
    ' Verificar que existe el directorio fuente
    If Not objFSO.FolderExists(srcPath) Then
        LogMessage "Error: No existe el directorio fuente de modulos: " & srcPath
        Exit Function
    End If
    
    ' Abrir Access
    Set objAccess = OpenAccess(dbPath, gPassword)
    If objAccess Is Nothing Then
        LogMessage "Error: No se pudo abrir la base de datos para rebuild"
        Exit Function
    End If
    
    ' Obtener lista de archivos de modulos
    moduleFiles = GetModuleFiles(srcPath, extensions, includeSubdirs, filePattern)
    
    If UBound(moduleFiles) >= 0 Then
        LogMessage "Encontrados " & (UBound(moduleFiles) + 1) & " archivos de modulos"
        
        ' Procesar cada archivo de modulo
        For i = 0 To UBound(moduleFiles)
            moduleFile = moduleFiles(i)
            moduleName = objFSO.GetBaseName(moduleFile)
            
            LogMessage "Procesando modulo: " & moduleName & " desde " & moduleFile
            
            ' Leer contenido del archivo
            moduleContent = ReadModuleFile(moduleFile)
            
            If moduleContent <> "" Then
                ' Importar o actualizar modulo en Access
                If ImportModuleToAccess(moduleName, moduleContent, moduleFile) Then
                Else
                    LogMessage "Error importando modulo: " & moduleName
                End If
            Else
                LogMessage "Error leyendo archivo de modulo: " & moduleFile
            End If
        Next
        
        RebuildModules = True
        LogMessage "Rebuild completado exitosamente"
    Else
        LogMessage "No se encontraron archivos de modulos en: " & srcPath
        RebuildModules = True ' No es error si no hay modulos
    End If
    
    ' Cerrar Access
    CloseAccess objAccess
    
    If Err.Number <> 0 Then
        LogMessage "Error durante rebuild: " & Err.Description
        Err.Clear
        RebuildModules = False
    End If
End Function

' Función para actualizar modulos VBA desde archivos fuente
Function UpdateModules(dbPath)
    On Error Resume Next
    
    Dim config, srcPath, extensions, includeSubdirs, filePattern
    Dim moduleFiles, i, moduleFile, moduleName, moduleContent
    Dim dbModules, needsUpdate
    
    UpdateModules = False
    
    ' Cargar configuración
    Set config = LoadConfig(gConfigPath)
    srcPath = config("MODULES_SrcPath")
    extensions = config("MODULES_Extensions")
    includeSubdirs = LCase(config("MODULES_IncludeSubdirectories")) = "true"
    filePattern = config("MODULES_FilePattern")
    
    LogMessage "Iniciando update de modulos VBA desde: " & srcPath
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
            
            ' Verificar si necesita actualización
            needsUpdate = ModuleNeedsUpdate(moduleFile, moduleName, dbModules)
            
            If needsUpdate Then
                LogMessage "Actualizando modulo: " & moduleName & " desde " & moduleFile
                
                ' Leer contenido del archivo
                moduleContent = ReadModuleFile(moduleFile)
                
                If moduleContent <> "" Then
                    ' Importar o actualizar modulo en Access
                    If ImportModuleToAccess(moduleName, moduleContent, moduleFile) Then
                    Else
                        LogMessage "Error actualizando modulo: " & moduleName
                    End If
                Else
                    LogMessage "Error leyendo archivo de modulo: " & moduleFile
                End If
            Else
                LogMessage "modulo " & moduleName & " está actualizado"
            End If
        Next
        
        UpdateModules = True
        LogMessage "Update completado exitosamente"
    Else
        LogMessage "No se encontraron archivos de modulos en: " & srcPath
        UpdateModules = True ' No es error si no hay modulos
    End If
    
    ' Cerrar Access
    CloseAccess
    
    If Err.Number <> 0 Then
        LogMessage "Error durante update: " & Err.Description
        Err.Clear
        UpdateModules = False
    End If
End Function

' Función para obtener lista de archivos de modulos
Function GetModuleFiles(srcPath, extensions, includeSubdirs, filePattern)
    Dim files(), fileCount, folder, file, subFolder
    Dim extArray, i, j, ext, fileName
    
    fileCount = 0
    ReDim files(-1)
    
    ' Convertir extensiones a array
    extArray = Split(extensions, ",")
    For i = 0 To UBound(extArray)
        extArray(i) = Trim(extArray(i))
    Next
    
    Set folder = objFSO.GetFolder(srcPath)
    
    ' Procesar archivos en el directorio actual
    For Each file In folder.Files
        fileName = file.Name
        
        ' Verificar extensión
        For j = 0 To UBound(extArray)
            ext = extArray(j)
            If Right(LCase(fileName), Len(ext)) = LCase(ext) Then
                ' Verificar patrón si se especifica
                If filePattern = "*" Or InStr(LCase(fileName), LCase(filePattern)) > 0 Then
                    ReDim Preserve files(fileCount)
                    files(fileCount) = file.Path
                    fileCount = fileCount + 1
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
    
    GetModuleFiles = files
End Function

' Función para leer contenido de archivo de modulo
Function ReadModuleFile(filePath)
    On Error Resume Next
    
    Dim file, content
    
    ReadModuleFile = ""
    
    If objFSO.FileExists(filePath) Then
        Set file = objFSO.OpenTextFile(filePath, 1, False, -1) ' Unicode
        If Err.Number <> 0 Then
            Err.Clear
            Set file = objFSO.OpenTextFile(filePath, 1, False, 0) ' ASCII
        End If
        
        If Not file.AtEndOfStream Then
            content = file.ReadAll
            ReadModuleFile = content
        End If
        
        file.Close
    End If
    
    If Err.Number <> 0 Then
        LogMessage "Error leyendo archivo: " & filePath & " - " & Err.Description
        Err.Clear
    End If
End Function

' Función para importar modulo a Access
Function ImportModuleToAccess(moduleName, moduleContent, filePath)
    On Error Resume Next
    
    Dim moduleType, fileExt
    
    ImportModuleToAccess = False
    
    ' Determinar tipo de modulo por extensión
    fileExt = LCase(objFSO.GetExtensionName(filePath))
    
    Select Case fileExt
        Case "bas"
            moduleType = acModule ' Modulo estandar
        Case "cls"
            moduleType = acClassModule ' Modulo de clase
        Case Else
            LogMessage "Extension de archivo no soportada: " & fileExt
            Exit Function
    End Select
    
    ' Intentar eliminar modulo existente
    DeleteExistingModule moduleName, moduleType
    If Err.Number <> 0 Then
        LogMessage "Advertencia eliminando modulo existente " & moduleName & ": " & Err.Description
        Err.Clear ' Continuar aunque no se pueda eliminar
    End If
    
    ' Importar nuevo modulo usando VBE
    If ImportVBAModuleSafe(moduleName, moduleContent, moduleType) Then
        ImportModuleToAccess = True
        LogMessage "Modulo " & moduleName & " importado exitosamente"
    Else
        LogMessage "Error importando modulo " & moduleName
        ImportModuleToAccess = False
    End If
End Function

' Función para eliminar modulo existente
Sub DeleteExistingModule(moduleName, moduleType)
    On Error Resume Next
    
    Select Case moduleType
        Case acModule, acClassModule
            ' Eliminar modulo VBA
            objAccess.VBE.VBProjects(0).VBComponents.Remove objAccess.VBE.VBProjects(0).VBComponents(moduleName)
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
    Set vbProject = objAccess.VBE.VBProjects(0)
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
    
    Set vbProject = objAccess.VBE.VBProjects(0)
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
        componentCount = objAccess.VBE.VBProjects(0).VBComponents.Count
        If componentCount > 0 Then
            For i = 1 To componentCount
                Set vbComp = objAccess.VBE.VBProjects(0).VBComponents(i)
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
    Set file = objFSO.CreateTextFile(filePath, True, True) ' Sobrescribir, Unicode
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
    gTestMode = False
    gDryRun = False
    gPassword = ""
    gOutputPath = gScriptDir & "\output"
    
    ' Cargar configuracion de modulos
    Set gConfig = LoadConfig(gConfigPath)
    
    ' Integrar configuración de base de datos desde cli.ini
    If gConfig("DATABASE_Password") <> "" Then
        gPassword = gConfig("DATABASE_Password")
    End If
    
    ' Configurar timeout desde cli.ini
    gTimeout = CInt(gConfig("DATABASE_Timeout"))
    
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
            gTestMode = True
            LogMessage "Modo de prueba activado"
            RunTests
            WScript.Quit 0
        ElseIf objArgs(i) = "/dry-run" Then
            gDryRun = True
            LogMessage "Modo simulacion activado"
        ElseIf objArgs(i) = "/validate" Then
            ValidateConfig
            WScript.Quit 0
        ElseIf objArgs(i) = "/verbose" Then
            gVerbose = True
            LogMessage "Modo verbose activado"
        ElseIf objArgs(i) = "/debug" Then
            LogMessage "Modo debug activado"
        ElseIf Left(objArgs(i), 10) = "/password:" Then
            gPassword = Mid(objArgs(i), 11)
            If gVerbose Then LogMessage "Password configurada desde parametro"
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
        Case "extract-all"
            If cleanArgCount >= 2 Then
                gDbPath = ResolvePath(cleanArgs(1))
                If cleanArgCount >= 3 Then 
                    gOutputPath = ResolvePath(cleanArgs(2))
                End If
                
                If Not gDryRun Then
                    ExtractAll gDbPath, gOutputPath
                Else
                    LogMessage "SIMULACION: Extraeria toda la informacion de " & gDbPath
                End If
            Else
                LogError "Faltan argumentos para extract-all"
                ShowHelp
            End If
            
        Case "extract-tables"
            If cleanArgCount >= 2 Then
                gDbPath = ResolvePath(cleanArgs(1))
                If cleanArgCount >= 3 Then 
                    gOutputPath = ResolvePath(cleanArgs(2))
                End If
                
                If Not gDryRun Then
                    Set objAccess = OpenAccess(gDbPath, gPassword)
                    If Not objAccess Is Nothing Then
                        ExtractTables gOutputPath
                        CloseAccess objAccess
                    End If
                Else
                    LogMessage "SIMULACION: Extraeria informacion de tablas de " & gDbPath
                End If
            Else
                LogError "Faltan argumentos para extract-tables"
                ShowHelp
            End If
            
        Case "extract-forms"
            If cleanArgCount >= 2 Then
                gDbPath = ResolvePath(cleanArgs(1))
                If cleanArgCount >= 3 Then 
                    gOutputPath = ResolvePath(cleanArgs(2))
                End If
                
                If Not gDryRun Then
                    Set objAccess = OpenAccess(gDbPath, gPassword)
                    If Not objAccess Is Nothing Then
                        ExtractForms gOutputPath
                        CloseAccess objAccess
                    End If
                Else
                    LogMessage "SIMULACION: Extraeria informacion de formularios de " & gDbPath
                End If
            Else
                LogError "Faltan argumentos para extract-forms"
                ShowHelp
            End If
            
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
            If cleanArgCount >= 2 Then
                gDbPath = ResolvePath(cleanArgs(1))
                
                If Not gDryRun Then
                    If gVerbose Then WScript.Echo "Ejecutando rebuild de modulos VBA..."
                    If Not RebuildModules(gDbPath) Then
                        WScript.Echo "Error: No se pudo completar el rebuild de modulos"
                        WScript.Quit 1
                    End If
                Else
                    WScript.Echo "[DRY-RUN] Se ejecutaria rebuild de modulos VBA en: " & gDbPath
                End If
            Else
                WScript.Echo "Error: Se requiere especificar la ruta de la base de datos"
                ShowHelp
                WScript.Quit 1
            End If
            
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
Sub RunTests()
    LogMessage "=== INICIANDO TESTS DEL CLI ==="
    
    ' Test 1: Verificar configuración
    TestLoadConfig
    
    ' Test 2: Verificar funciones utilitarias
    TestUtilityFunctions
    
    ' Test 3: Verificar conexión a base de datos (si existe archivo de prueba)
    TestDatabaseConnection
    
    LogMessage "=== TESTS COMPLETADOS ==="
End Sub

Sub TestLoadConfig()
    LogMessage "Test: Carga de configuración..."
    
    On Error Resume Next
    Dim testConfig
    Set testConfig = LoadConfig(gConfigPath)
    
    If Err.Number = 0 Then
        LogMessage "✓ Configuración cargada correctamente"
        LogVerbose "  - Elementos de configuración: " & testConfig.Count
    Else
        LogMessage "✗ Error cargando configuración: " & Err.Description
    End If
    
    On Error GoTo 0
End Sub

Sub TestUtilityFunctions()
    LogMessage "Test: Funciones utilitarias..."
    
    ' Test conversión de tipos de campo
    Dim fieldTypeName
    fieldTypeName = GetFieldTypeName(10) ' dbText
    If fieldTypeName = "Text" Then
        LogMessage "✓ GetFieldTypeName funciona correctamente"
    Else
        LogMessage "✗ GetFieldTypeName falló: " & fieldTypeName
    End If
    
    ' Test conversión de tipos de control
    Dim controlTypeName
    controlTypeName = GetControlTypeName(acTextBox)
    If controlTypeName = "Text Box" Then
        LogMessage "✓ GetControlTypeName funciona correctamente"
    Else
        LogMessage "✗ GetControlTypeName falló: " & controlTypeName
    End If
End Sub

Sub TestDatabaseConnection()
    LogMessage "Test: Conexión a base de datos..."
    
    If gDbPath = "" Then
        LogMessage "⚠ No hay ruta de BD configurada para test"
        Exit Sub
    End If
    
    If Not objFSO.FileExists(gDbPath) Then
        LogMessage "⚠ Archivo de BD no existe: " & gDbPath
        Exit Sub
    End If
    
    On Error Resume Next
    Set objAccess = OpenAccess(gDbPath, gPassword)
    If Not objAccess Is Nothing Then
        LogMessage "✓ Conexión a BD exitosa"
        LogVerbose "  - Archivo: " & objFSO.GetFileName(gDbPath)
        CloseAccess objAccess
    Else
        LogMessage "✗ Error conectando a BD: " & Err.Description
    End If
    
    On Error GoTo 0
End Sub

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
                If ImportModuleToAccess(objFSO.GetBaseName(moduleName), moduleContent, moduleFile) Then
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
                If ImportModuleToAccess(moduleName, moduleContent, moduleFile) Then
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
                    If ImportModuleToAccess(moduleName, moduleContent, moduleFile) Then
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
    
    CloseAccess objAccess
End Sub

' Función para validar configuración
Sub ValidateConfig()
    LogMessage "Validando configuración..."
    
    If objFSO.FileExists(gConfigPath) Then
        LogMessage "Archivo de configuración encontrado: " & gConfigPath
        Set objConfig = LoadConfig(gConfigPath)
        LogMessage "Configuración cargada correctamente"
        LogMessage "Elementos de configuración: " & objConfig.Count
    Else
        LogMessage "Archivo de configuración no encontrado: " & gConfigPath
        LogMessage "Se usarán valores por defecto"
    End If
End Sub

' Ejecutar función principal
Main()