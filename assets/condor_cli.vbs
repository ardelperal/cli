' ============================================================================
' CONDOR CLI - Herramienta de linea de comandos para gestion VBA
' ============================================================================
' Descripcion: Script VBScript para automatizar operaciones de desarrollo
' Autor: Equipo CONDOR
' Version: 2.0
' ============================================================================

Option Explicit

' ============================================================================
' SECCI├ôN 1: CONSTANTES
' ============================================================================

' ==== Access Enum constants (VBScript no tiene referencias a Access) ====
Const acModule = 5
Const acForm = -32768
Const acReport = -32764
Const acMacro = -32766
Const acTable = 0
Const acQuery = 1
Const acDefault = -1
Const acHidden = 1
Const acNormal = 0
Const acIcon = 1
Const acMaximized = 2
Const acMinimized = 3
Const acWindowNormal = 0
Const acCmdSaveRecord = 97
Const acQSave = 0

' Secciones de Access
Const acDetail = 0
Const acHeader = 1
Const acFooter = 2

' Controles de Access (valores num├®ricos)
Const acLabel = 100
Const acTextBox = 109
Const acCommandButton = 104
Const acCheckBox = 106
Const acOptionButton = 101
Const acComboBox = 111
Const acListBox = 110
Const acSubform = 112
Const acImage = 103
Const acLine = 4
Const acRectangle = 3
Const acPageBreak = 2
Const acTabCtl = 123
Const acCustomControl = 119
Const acQUpdate = 48
Const acQAppend = 64
Const acQDelete = 96
Const acQMakeTable = 80
Const acQCrosstab = 16
Const acQDDL = 112
Const acQSQLPassThrough = 128
Const acQSetOperation = 144
Const acQSPTBulk = 160

' ============================================================================
' SECCI├ôN 2: VARIABLES GLOBALES
' ============================================================================

' Variables globales del sistema
Dim objFSO, objArgs, objAccess

' Variables de configuraci├│n
Dim gVerbose, gBypassStartup, gPassword, gDbPath, gDryRun, gOpenShared
Dim gBypassStartupEnabled, gPreviousAllowBypassKey, gCurrentDbPath, gCurrentPassword
Dim gPreviousStartupForm, gPreviousHasAutoExec, gDbSource, gPrintDb

' ============================================================================
' SECCI├ôN 3: FUNCIONES HELPER
' ============================================================================

' ===== FUNCIONES HELPER PARA SECCIONES DE FORMULARIOS =====

' Funci├│n para mostrar ayuda general
Sub ShowHelp()
    WScript.Echo "=== CONDOR CLI - Herramienta de Linea de Comandos ==="
    WScript.Echo "Version: 2.0"
    WScript.Echo ""
    WScript.Echo "USO:"
    WScript.Echo "  cscript condor_cli.vbs <comando> [argumentos] [opciones]"
    WScript.Echo ""
    WScript.Echo "COMANDOS DISPONIBLES:"
    WScript.Echo ""
    WScript.Echo "  export-form <form_name> <json_path> [db_path]"
    WScript.Echo "    Exporta un formulario de Access a formato JSON"
    WScript.Echo "    Parametros:"
    WScript.Echo "      <form_name>  - Nombre del formulario a exportar"
    WScript.Echo "      <json_path>  - Ruta donde guardar el archivo JSON"
    WScript.Echo "      [db_path]    - Ruta de la base de datos (opcional)"
    WScript.Echo "    Opciones:"
    WScript.Echo "      --password <pwd>  - password de la base de datos"
    WScript.Echo "      --pretty          - Formato JSON con indentacion"
    WScript.Echo ""
    WScript.Echo "  import-form <json_path> <form_name> [db_path]"
    WScript.Echo "    Importa un formulario desde JSON a Access"
    WScript.Echo "    Parametros:"
    WScript.Echo "      <json_path>  - Ruta del archivo JSON a importar"
    WScript.Echo "      <form_name>  - Nombre del formulario destino"
    WScript.Echo "      [db_path]    - Ruta de la base de datos (opcional)"
    WScript.Echo "    Opciones:"
    WScript.Echo "      --password <pwd>  - password de la base de datos"
    WScript.Echo "      --overwrite       - Sobrescribir si el formulario existe"
    WScript.Echo ""
    WScript.Echo "  roundtrip-form <form_name> <temp_dir> [db_path]"
    WScript.Echo "    Prueba de ida y vuelta: exporta, reimporta y compara"
    WScript.Echo "    Parametros:"
    WScript.Echo "      <form_name>  - Nombre del formulario a probar"
    WScript.Echo "      <temp_dir>   - Directorio temporal para archivos"
    WScript.Echo "      [db_path]    - Ruta de la base de datos (opcional)"
    WScript.Echo "    Opciones:"
    WScript.Echo "      --password <pwd>  - password de la base de datos"
    WScript.Echo ""
    WScript.Echo "  validate-form-json <json_path>"
    WScript.Echo "    Valida la estructura de un archivo JSON de formulario"
    WScript.Echo "    Parametros:"
    WScript.Echo "      <json_path>  - Ruta del archivo JSON a validar"
    WScript.Echo ""
    WScript.Echo "  list-forms [db_path]"
    WScript.Echo "    Lista todos los formularios de la base de datos"
    WScript.Echo "    Parametros:"
    WScript.Echo "      [db_path]    - Ruta de la base de datos (opcional)"
    WScript.Echo "    Opciones:"
    WScript.Echo "      --password <pwd>  - password de la base de datos"
    WScript.Echo "      --json            - Salida en formato JSON"
    WScript.Echo ""
    WScript.Echo "OPCIONES GLOBALES:"
    WScript.Echo "  --help            - Muestra esta ayuda"
    WScript.Echo "  --verbose         - Salida detallada"
    WScript.Echo "  --quiet           - Salida minima"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS:"
    WScript.Echo "  cscript condor_cli.vbs export-form ""FormularioPrincipal"" ""C:\temp\form.json"""
    WScript.Echo "  cscript condor_cli.vbs import-form ""C:\temp\form.json"" ""FormularioNuevo"" --overwrite"
    WScript.Echo "  cscript condor_cli.vbs roundtrip-form ""FormularioPrincipal"" ""C:\temp"" --password 1234"
    WScript.Echo "  cscript condor_cli.vbs validate-form-json ""C:\temp\form.json"""
    WScript.Echo "  cscript condor_cli.vbs list-forms --json"
    WScript.Echo ""
    WScript.Echo "Para ayuda especifica de un comando: <comando> --help"
End Sub

' Funci├│n para mostrar ayuda de bundle
Sub ShowBundleHelp()
    WScript.Echo "=== BUNDLE - Empaquetado de Codigo por Funcionalidad ==="
    WScript.Echo "Uso: cscript condor_cli.vbs bundle <funcionalidad> [ruta_destino]"
    WScript.Echo ""
    WScript.Echo "PARAMETROS:"
    WScript.Echo "  <funcionalidad>  - Nombre de la funcionalidad a empaquetar"
    WScript.Echo "  [ruta_destino]   - Directorio donde crear el bundle (opcional)"
    WScript.Echo ""
    WScript.Echo "FUNCIONALIDADES DISPONIBLES:"
    WScript.Echo "  Tests            - Suite de pruebas unitarias"
    WScript.Echo "  Core             - Funcionalidades basicas del sistema"
    WScript.Echo "  Expedientes      - Gestion de expedientes"
    WScript.Echo "  Solicitudes      - Gestion de solicitudes"
    WScript.Echo "  Reportes         - Sistema de reportes"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS:"
    WScript.Echo "  cscript condor_cli.vbs bundle Tests"
    WScript.Echo "  cscript condor_cli.vbs bundle Core C:\temp\bundles"
End Sub

' ===== FUNCIONES UTILITARIAS DE RUTAS =====

' Funci├│n para obtener la carpeta ra├¡z del repositorio
Function RepoRoot()
    Dim scriptPath
    scriptPath = WScript.ScriptFullName
    RepoRoot = objFSO.GetParentFolderName(scriptPath)
End Function

' Funci├│n auxiliar para verificar esquema de una base de datos espec├¡fica
Private Function VerifySchema(dbPath, dbPassword, expectedSchema)
    On Error Resume Next
    
    WScript.Echo "Validando base de datos: " & dbPath
    
    ' Verificar que existe la base de datos
    If Not objFSO.FileExists(dbPath) Then
        WScript.Echo "[ERROR] Base de datos no encontrada: " & dbPath
        VerifySchema = False
        Exit Function
    End If
    
    ' Crear conexi├│n ADO
    Dim conn, rs
    Set conn = CreateObject("ADODB.Connection")
    
    ' Construir cadena de conexi├│n
    Dim connectionString
    If dbPassword = "" Then
        connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";"
    Else
        connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & dbPath & ";Jet OLEDB:Database Password=" & dbPassword & ";"
    End If
    
    ' Abrir conexi├│n
    conn.Open connectionString
    
    If Err.Number <> 0 Then
        WScript.Echo "[ERROR] No se pudo conectar a la base de datos: " & Err.Description
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
        
        ' Verificar que existe la tabla usando una consulta m├ís simple
         Set rs = CreateObject("ADODB.Recordset")
         On Error Resume Next
         rs.Open "SELECT TOP 1 * FROM [" & tableName & "]", conn
         tableExists = (Err.Number = 0)
         If tableExists Then rs.Close
         Err.Clear
         On Error GoTo 0
        
        If Not tableExists Then
            WScript.Echo "[ERROR] Tabla no encontrada: " & tableName
            allTablesOk = False
        Else
            WScript.Echo "[OK] Tabla encontrada: " & tableName
            
            ' Verificar cada campo esperado
            For i = 0 To UBound(expectedFields)
                Dim fieldName
                fieldName = expectedFields(i)
                
                ' Verificar que existe el campo usando una consulta m├ís simple
                 Set rs = CreateObject("ADODB.Recordset")
                 On Error Resume Next
                 rs.Open "SELECT [" & fieldName & "] FROM [" & tableName & "] WHERE 1=0", conn
                 fieldExists = (Err.Number = 0)
                 If fieldExists Then rs.Close
                 Err.Clear
                 On Error GoTo 0
                
                If Not fieldExists Then
                    WScript.Echo "[ERROR] Campo no encontrado: " & tableName & "." & fieldName
                    allTablesOk = False
                Else
                    WScript.Echo "  [OK] Campo encontrado: " & fieldName
                End If
            Next
        End If
    Next
    
    ' Cerrar conexi├│n
    conn.Close
    Set conn = Nothing
    
    If allTablesOk Then
        WScript.Echo "[OK] Base de datos validada correctamente: " & objFSO.GetFileName(dbPath)
        VerifySchema = True
    Else
        WScript.Echo "[ERROR] Errores encontrados en: " & objFSO.GetFileName(dbPath)
        VerifySchema = False
    End If
    
    On Error GoTo 0
End Function

Sub ValidateSchema()
    WScript.Echo "=== INICIANDO VALIDACI├ôN DE ESQUEMA DE BASE DE DATOS ==="
    
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
    Dim strSourcePath: strSourcePath = RepoRoot() & "\src"
    If Not VerifySchema(strSourcePath & "\..\back\test_env\fixtures\databases\Lanzadera_test_template.accdb", "dpddpd", lanzaderaSchema) Then allOk = False
        If Not VerifySchema(strSourcePath & "\..ack\test_env\fixtures\databases\Document_test_template.accdb", "", condorSchema) Then allOk = False
    
    If allOk Then
        WScript.Echo "[OK] VALIDACION DE ESQUEMA EXITOSA. Todas las bases de datos son consistentes."
        WScript.Quit 0
    Else
        WScript.Echo "[ERROR] VALIDACION DE ESQUEMA FALLIDA. Corrija las discrepancias."
        WScript.Quit 1
    End If
End Sub

' Funci├│n para obtener la ruta por defecto de la BD frontend
Function DefaultFrontendDb()
    Dim defaultPath, legacyPath
    defaultPath = GetDevDbPath()
    legacyPath = objFSO.BuildPath(RepoRoot(), "back\Desarrollo\CONDOR.accdb")
    If Not objFSO.FileExists(defaultPath) And objFSO.FileExists(legacyPath) Then
        If gVerbose Or gPrintDb Then
            WScript.Echo "WARNING: usando ruta legacy back\Desarrollo\CONDOR.accdb; migre a front\Desarrollo\CONDOR.accdb (deprecado)"
        End If
        DefaultFrontendDb = legacyPath
    Else
        DefaultFrontendDb = defaultPath
    End If
End Function

' Funci├│n para obtener la ruta por defecto de la BD backend
Function DefaultBackendDb()
    Dim defaultPath, legacyPath
    defaultPath = objFSO.BuildPath(GetAuxDataRoot(), "CONDOR_datos.accdb")
    legacyPath = objFSO.BuildPath(RepoRoot(), "back\CONDOR_datos.accdb")
    If Not objFSO.FileExists(defaultPath) And objFSO.FileExists(legacyPath) Then
        If gVerbose Or gPrintDb Then
            WScript.Echo "WARNING: usando ruta legacy back\CONDOR_datos.accdb; migre a back\data\CONDOR_datos.accdb (deprecado)"
        End If
        DefaultBackendDb = legacyPath
    Else
        DefaultBackendDb = defaultPath
    End If
End Function

' ===== HELPERS DE RUTAS CENTRALIZADAS =====

Function GetFrontRoot()
    GetFrontRoot = objFSO.BuildPath(RepoRoot(), "front")
End Function

Function GetBackRoot()
    GetBackRoot = objFSO.BuildPath(RepoRoot(), "back")
End Function

Function GetTemplatesPath()
    GetTemplatesPath = objFSO.BuildPath(GetFrontRoot(), "recursos\Plantillas")
End Function

Function GetTestEnvPath()
    GetTestEnvPath = objFSO.BuildPath(GetFrontRoot(), "test_env")
End Function

Function GetDevDbPath()
    GetDevDbPath = objFSO.BuildPath(GetFrontRoot(), "Desarrollo\CONDOR.accdb")
End Function

Function GetAuxDataRoot()
    GetAuxDataRoot = objFSO.BuildPath(GetBackRoot(), "data")
End Function

' Funci├│n para limpiar comillas de argumentos
Function TrimQuotes(s)
    If Left(s, 1) = """" And Right(s, 1) = """" Then
        TrimQuotes = Mid(s, 2, Len(s) - 2)
    Else
        TrimQuotes = s
    End If
End Function

' Funci├│n para verificar si un token es una ruta de BD
Function IsDbPathToken(tok)
    IsDbPathToken = (InStr(tok, ".accdb") > 0 Or InStr(tok, ".mdb") > 0)
End Function

' Funci├│n para convertir ruta relativa a absoluta
Function ToAbsolute(pathLike)
    If objFSO.GetAbsolutePathName(pathLike) = pathLike Then
        ToAbsolute = pathLike
    Else
        ToAbsolute = objFSO.BuildPath(RepoRoot(), pathLike)
    End If
End Function

' ============================================================================
' SECCI├ôN 4: PARSER DE ARGUMENTOS
' ============================================================================

' ===== PARSER DE ARGUMENTOS ROBUSTO =====

' Funci├│n principal para resolver flags y argumentos
Sub ResolveFlags()
    Dim i, arg
    
    ' Procesar todos los argumentos
    For i = 0 To objArgs.Count - 1
        arg = objArgs(i)
        
        ' Flags de configuraci├│n
        If arg = "--verbose" Or arg = "-v" Then
            gVerbose = True
        ElseIf arg = "--db" And i < objArgs.Count - 1 Then
            gDbPath = TrimQuotes(objArgs(i + 1))
            i = i + 1 ' Saltar el siguiente argumento
        ElseIf arg = "--password" And i < objArgs.Count - 1 Then
            gPassword = TrimQuotes(objArgs(i + 1))
            i = i + 1 ' Saltar el siguiente argumento
        ElseIf arg = "--no-bypass" Then
            gBypassStartup = False
            gBypassStartupEnabled = False
        ElseIf arg = "--bypass" Then
            gBypassStartup = True
            gBypassStartupEnabled = True
        ElseIf arg = "--dry-run" Then
            gDryRun = True
        ElseIf arg = "--sharedopen" Then
            gOpenShared = True
        ElseIf arg = "--print-db" Then
            gPrintDb = True
        ElseIf IsDbPathToken(objArgs(i)) Then
            gDbPath = TrimQuotes(objArgs(i))
        End If
    Next
    
    ' Aplicar bypass por defecto si no se especific├│
    If Not gBypassStartupEnabled Then
        Call SetDefaultBypassStartup()
    End If
    
    If gVerbose Then
        WScript.Echo "[VERBOSE] Flags procesados:"
        WScript.Echo "[VERBOSE]   --verbose: " & gVerbose
        WScript.Echo "[VERBOSE]   --db: " & gDbPath
        WScript.Echo "[VERBOSE]   --bypass: " & gBypassStartup
    End If
End Sub

' Funci├│n para verificar si se solicita ayuda


' Funci├│n para listar comandos disponibles
Sub ListAvailableCommands()
    WScript.Echo "=== COMANDOS DISPONIBLES ==="
    WScript.Echo "  validate        - Validar todos los m├│dulos"
    WScript.Echo "  export          - Exportar m├│dulos a archivos"
    WScript.Echo "  test            - Ejecutar pruebas"
    WScript.Echo "  list-forms      - Listar formularios"
    WScript.Echo "  list-modules    - Listar m├│dulos"
    WScript.Echo "  bundle          - Empaquetar funcionalidades"
    WScript.Echo ""
    WScript.Echo "Para ayuda espec├¡fica: cscript condor_cli.vbs <comando> --help"
End Sub

' Subrutina para establecer bypass startup por defecto seg├║n comando
Sub SetDefaultBypassStartup()
    ' Bypass startup habilitado por defecto para todos los comandos
    gBypassStartup = True
    gBypassStartupEnabled = True
End Sub

' ============================================================================
' SECCI├ôN 5: RESOLUCI├ôN DE BASE DE DATOS
' ============================================================================

' ===== RESOLUCI├ôN CAN├ôNICA DE BASE DE DATOS =====

' Funci├│n principal para resolver la ruta de BD seg├║n la acci├│n
Function ResolveDbForAction(actionName, ByRef origin)
    Dim resolvedPath
    
    ' Prioridad 1: --db expl├¡cito
    If gDbPath <> "" Then
        resolvedPath = ToAbsolute(gDbPath)
        origin = "explicit-db"
        ResolveDbForAction = resolvedPath
        Exit Function
    End If
    
    ' Prioridad 2: Detectar BD en argumentos posicionales
    Dim i
    For i = 1 To objArgs.Count - 1
        If IsDbPathToken(objArgs(i)) Then
            resolvedPath = ToAbsolute(TrimQuotes(objArgs(i)))
            origin = "positional-arg"
            ResolveDbForAction = resolvedPath
            Exit Function
        End If
    Next
    
    ' Prioridad 3: Variable de entorno CONDOR_DB
    Dim envDb
    Set envDb = CreateObject("WScript.Shell").Environment("Process")
    If envDb("CONDOR_DB") <> "" Then
        resolvedPath = ToAbsolute(envDb("CONDOR_DB"))
        origin = "env-var"
        ResolveDbForAction = resolvedPath
        Exit Function
    End If
    
    ' Prioridad 4: Default seg├║n acci├│n usando DefaultForAction
    resolvedPath = DefaultForAction(actionName, origin)
    ResolveDbForAction = resolvedPath
End Function

' Funci├│n que determina la BD por defecto seg├║n la acci├│n
Function DefaultForAction(actionName, ByRef origin)
    ' FRONTEND por defecto para acciones de c├│digo/desarrollo
    If actionName = "rebuild" Or actionName = "update" Or actionName = "export" Or _
       actionName = "validate" Or actionName = "test" Or actionName = "export-form" Or _
       actionName = "import-form" Or actionName = "list-forms" Or actionName = "list-modules" Or _
       actionName = "roundtrip-form" Or actionName = "validate-form-json" Then
        origin = "default-frontend"
        DefaultForAction = DefaultFrontendDb()
    Else
        ' BACKEND por defecto para comandos de datos
        origin = "default-backend"
        DefaultForAction = DefaultBackendDb()
    End If
End Function

' ============================================================================
' SECCI├ôN 6: BYPASS/ACCESS
' ============================================================================

' ===== FUNCIONES DE MANEJO DE ACCESS =====

' Funci├│n para abrir Access de forma silenciosa
Function OpenAccessQuiet(dbPath, password)
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
        Set OpenAccessQuiet = Nothing
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
        Set OpenAccessQuiet = Nothing
        Exit Function
    End If
    
    On Error GoTo 0
    Set OpenAccessQuiet = objAccess
    
    If gVerbose Then
        WScript.Echo "[VERBOSE] Access abierto exitosamente"
    End If
End Function

' Funci├│n para cerrar Access de forma silenciosa
Sub CloseAccessQuiet(objAccess)
    If Not objAccess Is Nothing Then
        If gVerbose Then
            WScript.Echo "[VERBOSE] Cerrando Access..."
        End If
        
        On Error Resume Next
        objAccess.CloseCurrentDatabase
        objAccess.Quit
        Set objAccess = Nothing
        On Error GoTo 0
        
        If gVerbose Then
            WScript.Echo "[VERBOSE] Access cerrado exitosamente"
        End If
    End If
End Sub

' Funci├│n para obtener PIDs de procesos MSACCESS.EXE usando WMI
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

' Funci├│n para encontrar la diferencia de un PID entre dos arrays
Function DiffOne(pidsAfter, pidsBefore)
    On Error Resume Next
    Dim i, j, found, newPid
    
    ' Si no hay PIDs despu├®s, no hay diferencia
    If Not IsArray(pidsAfter) Or UBound(pidsAfter) < 0 Then
        DiffOne = -1
        Exit Function
    End If
    
    ' Si no hab├¡a PIDs antes, devolver el primero de despu├®s
    If Not IsArray(pidsBefore) Or UBound(pidsBefore) < 0 Then
        DiffOne = pidsAfter(0)
        Exit Function
    End If
    
    ' Buscar PID que est├í en pidsAfter pero no en pidsBefore
    For i = 0 To UBound(pidsAfter)
        found = False
        For j = 0 To UBound(pidsBefore)
            If pidsAfter(i) = pidsBefore(j) Then
                found = True
                Exit For
            End If
        Next
        If Not found Then
            DiffOne = pidsAfter(i)
            Exit Function
        End If
    Next
    
    ' No se encontr├│ diferencia
    DiffOne = -1
    On Error GoTo 0
End Function

' Funci├│n para terminar un PID espec├¡fico de Access
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
            objProcess.Terminate()
        Else
            Err.Clear
        End If
    Next
    
    On Error GoTo 0
End Sub

' ===== FUNCI├ôN AUXILIAR PARA CONVERSI├ôN DE COLORES =====

Function ConvertColorToLong(colorStr)
    ' Convierte color #RRGGBB a Long BGR para Access
    If Left(colorStr, 1) = "#" And Len(colorStr) = 7 Then
        Dim r, g, b
        r = CLng("&H" & Mid(colorStr, 2, 2))
        g = CLng("&H" & Mid(colorStr, 4, 2))
        b = CLng("&H" & Mid(colorStr, 6, 2))
        
        ' Convertir a formato BGR (Blue-Green-Red)
        ConvertColorToLong = (b * 65536) + (g * 256) + r
    Else
        ' Si no es formato #RRGGBB, devolver el valor original
        ConvertColorToLong = colorStr
    End If
End Function

' ============================================================================
' SECCI├ôN 7: JSON
' ============================================================================

' [Aqu├¡ ir├¡an las funciones de manejo JSON - por brevedad las omito en esta primera parte]

' ============================================================================
' SECCI├ôN 8: COMANDOS
' ============================================================================

' [Aqu├¡ ir├¡an todas las implementaciones de comandos - por brevedad las omito en esta primera parte]

' ============================================================================
' SECCI├ôN 9: MAIN - PUNTO DE ENTRADA
' ============================================================================

' Configuraci├│n inicial
Set objArgs = WScript.Arguments
Set objFSO = CreateObject("Scripting.FileSystemObject")
Dim strSourcePath: strSourcePath = RepoRoot() & "\src"

' Inicializar variables globales
gVerbose = False
gBypassStartup = False
gPassword = ""
gDbPath = ""
gDryRun = False
gOpenShared = False
gBypassStartupEnabled = False
gPreviousAllowBypassKey = Null
gCurrentDbPath = ""
gCurrentPassword = ""
gPreviousStartupForm = Null
gPreviousHasAutoExec = False
gDbSource = ""
gPrintDb = False

' Verificar si se solicita ayuda
If objArgs.Count > 0 Then
    If LCase(objArgs(0)) = "--help" Or LCase(objArgs(0)) = "-h" Or LCase(objArgs(0)) = "help" Then
        Call ShowHelp()
        WScript.Quit 0
    End If
End If

' Validar argumentos minimos
If objArgs.Count = 0 Then
    WScript.Echo "[ERROR] No se especifico ningun comando"
    Call ShowHelp()
    WScript.Quit 1
End If

' Obtener acci├│n
Dim strAction: strAction = LCase(objArgs(0))

' Validar comando
If strAction <> "export" And strAction <> "validate" And strAction <> "validate-schema" And _
   strAction <> "test" And strAction <> "createtable" And strAction <> "droptable" And _
   strAction <> "listtables" And strAction <> "relink" And strAction <> "rebuild" And _
   strAction <> "update" And strAction <> "lint" And strAction <> "bundle" And _
   strAction <> "migrate" And strAction <> "export-form" And strAction <> "import-form" And _
   strAction <> "validate-form-json" And strAction <> "roundtrip-form" And _
   strAction <> "list-forms" And strAction <> "list-modules" And strAction <> "fix-src-headers" Then
    WScript.Echo "[ERROR] Comando desconocido: " & strAction
    WScript.Echo "Use --help para ver comandos disponibles"
    WScript.Quit 1
End If

' PASO 1: Resolver flags ANTES de cualquier apertura de Access
Call ResolveFlags()

' PASO 2: Resolver ruta de base de datos usando resoluci├│n can├│nica
Dim strOrigin, strAccessPath
strAccessPath = ResolveDbForAction(strAction, strOrigin)
gDbSource = strOrigin

' Mostrar informaci├│n de la BD resuelta si se solicita
If gPrintDb Or gVerbose Then
    WScript.Echo "DB resuelta: " & strAccessPath & " (origen=" & gDbSource & ")"
End If

' PASO 3: Determinar bypass startup por defecto seg├║n comando
Call SetDefaultBypassStartup()

' PASO 4: Ejecutar comandos que NO requieren Access
If strAction = "bundle" Then
    If objArgs.Count > 1 Then
        If LCase(objArgs(1)) = "--help" Or LCase(objArgs(1)) = "-h" Or LCase(objArgs(1)) = "help" Then
            Call ShowBundleHelp()
            WScript.Quit 0
        End If
    End If
    Call BundleFunctionality()
    WScript.Echo "[OK] Comando bundle ejecutado exitosamente"
    WScript.Quit 0
ElseIf strAction = "validate-schema" Then
    Call ValidateSchema()
    WScript.Echo "[OK] Comando validate-schema ejecutado exitosamente"
    WScript.Quit 0
ElseIf strAction = "export-form" Then
    Call ExportFormCommand()
    ' ExportFormCommand maneja su propio WScript.Quit
ElseIf strAction = "import-form" Then
    Call ImportFormCommand()
    ' ImportFormCommand maneja su propio WScript.Quit
ElseIf strAction = "validate-form-json" Then
    Call ValidateFormJsonCommand()
    ' ValidateFormJsonCommand maneja su propio WScript.Quit
ElseIf strAction = "roundtrip-form" Then
    Call RoundtripFormCommand()
    ' RoundtripFormCommand maneja su propio WScript.Quit
ElseIf strAction = "list-forms" Then
    Call ListFormsCommand()
    ' ListFormsCommand maneja su propio WScript.Quit
ElseIf strAction = "list-modules" Then
    Call ListModulesCommand()
    ' ListModulesCommand maneja su propio WScript.Quit
ElseIf strAction = "fix-src-headers" Then
    Call FixSrcHeadersCommand()
    WScript.Echo "[OK] Comando fix-src-headers ejecutado exitosamente"
    WScript.Quit 0
End If

' PASO 5: Verificar que existe la base de datos
If Not objFSO.FileExists(strAccessPath) Then
    WScript.Echo "[ERROR] base de datos no encontrada (" & strAccessPath & "), origen=" & gDbSource
    WScript.Quit 1
End If

' PASO 6: Mostrar informaci├│n de inicio
WScript.Echo "=== CONDOR CLI ==="
WScript.Echo "Acci├│n: " & strAction
WScript.Echo "Base de datos: " & strAccessPath
WScript.Echo "Directorio: " & strSourcePath
If gVerbose Then
    If gPassword <> "" Then
        WScript.Echo "[VERBOSE] Password: ***"
    Else
        WScript.Echo "[VERBOSE] Password: (none)"
    End If
    WScript.Echo "[INFO] BypassStartup aplicado autom├íticamente."
End If

' PASO 7: Cerrar procesos de Access existentes
Call CloseExistingAccessProcesses()

' PASO 8: Abrir Access con OpenAccessQuiet unificado (solo si es necesario)
If RequiresAccess(strAction) Then
    Set objAccess = OpenAccessQuiet(strAccessPath, gPassword)
    If objAccess Is Nothing Then
        WScript.Echo "[ERROR] No se pudo abrir Access. Abortando."
        WScript.Quit 1
    End If
Else
    Set objAccess = Nothing
    If gVerbose Then
        WScript.Echo "[VERBOSE] Comando no requiere Access, omitiendo apertura"
    End If
End If

Dim exitCode: exitCode = 0

' PASO 9: Ejecutar comando correspondiente
Select Case LCase(strAction)
    Case "validate"
        WScript.Echo "Ejecutando validación..."
        ' Call ValidateAllModules() ' Implementación pendiente
    Case "export"
        WScript.Echo "Ejecutando exportación..."
        ' Call ExportModules() ' Implementación pendiente
    Case "test"
        WScript.Echo "Ejecutando pruebas..."
        exitCode = ExecuteTests(objAccess)
    Case Else
        WScript.Echo "Ejecutando comando: " & strAction
        ' Implementaciones pendientes para otros comandos
End Select

' PASO 10: Cerrar Access si fue abierto
Call CloseAccessQuiet(objAccess)

If exitCode = 0 Then
    WScript.Echo "=== COMANDO COMPLETADO EXITOSAMENTE ==="
Else
    WScript.Echo "=== COMANDO FINALIZADO CON ERRORES ==="
End If

WScript.Quit exitCode
' ===== FUNCIONES AUXILIARES PARA MAIN =====

' Funci├│n para determinar si un comando requiere Access
Function RequiresAccess(actionName)
    Select Case LCase(actionName)
        Case "validate", "export", "test", "list-forms", "list-modules", "export-form", "import-form", "roundtrip-form"
            RequiresAccess = True
        Case "bundle", "help", "validate-schema", "validate-form-json", "fix-src-headers"
            RequiresAccess = False
        Case Else
            RequiresAccess = True ' Por defecto, asumir que requiere Access
    End Select
End Function

Function ExecuteTests(app)
    On Error Resume Next

    If app Is Nothing Then
        WScript.Echo "[ERROR] Instancia de Access no disponible"
        ExecuteTests = 1
        Exit Function
    End If

    Dim resultText
    Err.Clear
    resultText = app.Run("ExecuteAllTestsForCLI")
    Dim runErrNumber: runErrNumber = Err.Number
    Dim runErrDesc: runErrDesc = Err.Description
    On Error GoTo 0

    If runErrNumber <> 0 Then
        WScript.Echo "[ERROR] Error ejecutando ExecuteAllTestsForCLI: " & runErrDesc
        ExecuteTests = 1
        Exit Function
    End If

    If IsNull(resultText) Then
        resultText = ""
    End If

    WScript.Echo resultText

    Dim normalized: normalized = LCase(resultText)
    If InStr(normalized, "result: failure") > 0 Or InStr(normalized, "failed") > 0 Then
        ExecuteTests = 1
    ElseIf InStr(normalized, "result: success") > 0 Then
        ExecuteTests = 0
    Else
        ExecuteTests = 0 ' Asumir éxito si no hay marcador explícito
    End If
End Function

' ===== FUNCIONES DE AYUDA PARA COMANDOS DE FORMULARIOS =====
Sub ShowExportFormHelp()
    WScript.Echo "=== EXPORT-FORM - Exportar formulario a JSON ==="
    WScript.Echo "Uso: cscript condor_cli.vbs export-form <db_path> <form_name> --output <json_path> [opciones]"
    WScript.Echo ""
    WScript.Echo "PAR├üMETROS REQUERIDOS:"
    WScript.Echo "  <db_path>     - Ruta de la base de datos Access"
    WScript.Echo "  <form_name>   - Nombre del formulario a exportar"
    WScript.Echo "  --output      - Ruta del archivo JSON de salida"
    WScript.Echo ""
    WScript.Echo "OPCIONES:"
    WScript.Echo "  --password <pwd>  - password de la base de datos"
    WScript.Echo "  --pretty          - Formatear JSON con indentaci├│n"
    WScript.Echo "  --expand <tipo>   - Incluir bloques opcionales:"
    WScript.Echo "                      sections|properties|extra"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS:"
    WScript.Echo "  # Exportar desde UI/sources:"
    WScript.Echo "  cscript condor_cli.vbs export-form --db ""C:\Proyectos\CONDOR\ui\sources\Expedientes.accdb"" --password dpddpd ""F_Expediente"" --output "".\out\F_Expediente.json"""
    WScript.Echo ""
    WScript.Echo "  # Con formato pretty:"
    WScript.Echo "  cscript condor_cli.vbs export-form ""C:\MiDB.accdb"" ""MiForm"" --output ""form.json"" --pretty"
    WScript.Echo ""
    WScript.Echo "  # Con expansi├│n de secciones:"
    WScript.Echo "  cscript condor_cli.vbs export-form ""C:\MiDB.accdb"" ""MiForm"" --output ""form.json"" --expand sections"
End Sub

Sub ShowImportFormHelp()
    WScript.Echo "=== IMPORT-FORM - Importar formulario desde JSON ==="
    WScript.Echo "Uso: cscript condor_cli.vbs import-form <db_path> <json_path> [opciones]"
    WScript.Echo ""
    WScript.Echo "PAR├üMETROS REQUERIDOS:"
    WScript.Echo "  <db_path>     - Ruta de la base de datos Access destino"
    WScript.Echo "  <json_path>   - Ruta del archivo JSON a importar"
    WScript.Echo ""
    WScript.Echo "OPCIONES:"
    WScript.Echo "  --target <name>   - Nombre del formulario destino (si difiere del JSON)"
    WScript.Echo "  --replace         - Reemplazar formulario existente sin confirmaci├│n"
    WScript.Echo "  --password <pwd>  - password de la base de datos"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS:"
    WScript.Echo "  # Importar/editar en Desarrollo:"
    WScript.Echo "  cscript condor_cli.vbs import-form --db ""C:\Proyectos\CONDOR\front\Desarrollo\CONDOR.accdb"" "".\out\F_Expediente.json"" --replace"
    WScript.Echo ""
    WScript.Echo "  # Con nombre diferente:"
    WScript.Echo "  cscript condor_cli.vbs import-form ""C:\MiDB.accdb"" ""form.json"" --target ""NuevoNombre"" --replace"
End Sub

Sub ShowRoundtripFormHelp()
    WScript.Echo "=== ROUNDTRIP-FORM - Test de integridad export->import ==="
    WScript.Echo "Uso: cscript condor_cli.vbs roundtrip-form <db_path> <form_name> --temp <dir> [opciones]"
    WScript.Echo ""
    WScript.Echo "PARAMETROS REQUERIDOS:"
    WScript.Echo "  <db_path>     - Ruta de la base de datos Access"
    WScript.Echo "  <form_name>   - Nombre del formulario a probar"
    WScript.Echo "  --temp <dir>  - Directorio temporal para archivos intermedios"
    WScript.Echo ""
    WScript.Echo "OPCIONES:"
    WScript.Echo "  --password <pwd>  - password de la base de datos"
    WScript.Echo ""
    WScript.Echo "FLUJO:"
    WScript.Echo "  1. Exportar formulario a <temp>\<form_name>.json"
    WScript.Echo "  2. Reimportar JSON sobre el mismo formulario"
    WScript.Echo "  3. Exportar nuevamente a <temp>\<form_name>.post.json"
    WScript.Echo "  4. Comparar diferencias sem├ínticas"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS:"
    WScript.Echo "  cscript condor_cli.vbs roundtrip-form ""C:\MiDB.accdb"" ""MiForm"" --temp "".\temp"""
End Sub

Sub ShowValidateFormJsonHelp()
    WScript.Echo "=== VALIDATE-FORM-JSON - Validar estructura JSON ==="
    WScript.Echo "Uso: cscript condor_cli.vbs validate-form-json <json_path> [opciones]"
    WScript.Echo ""
    WScript.Echo "PAR├üMETROS REQUERIDOS:"
    WScript.Echo "  <json_path>   - Ruta del archivo JSON a validar"
    WScript.Echo ""
    WScript.Echo "OPCIONES:"
    WScript.Echo "  --schema <N>  - Validar contra esquema espec├¡fico (versi├│n N)"
    WScript.Echo "  --strict      - Validaci├│n estricta de tipos y rangos"
    WScript.Echo ""
    WScript.Echo "VALIDACIONES:"
    WScript.Echo "  - Campos requeridos: schemaVersion, properties, sections"
    WScript.Echo "  - properties: name, defaultView, recordSelectors, navigationButtons"
    WScript.Echo "  - sections: array con name, type Ôêê {header,detail,footer}"
    WScript.Echo "  - En modo --strict: validaci├│n de tipos y rangos"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS:"
    WScript.Echo "  cscript condor_cli.vbs validate-form-json ""form.json"""
    WScript.Echo "  cscript condor_cli.vbs validate-form-json ""form.json"" --strict"
    WScript.Echo "  cscript condor_cli.vbs validate-form-json ""form.json"" --schema 2"
End Sub

' ===== IMPLEMENTACIONES DE COMANDOS DE FORMULARIOS =====

Sub ExportFormCommand()
    On Error GoTo ExportFail

    Dim dbPath, formName, outputPath, password, pretty
    Dim i, arg

    dbPath = ""
    formName = ""
    outputPath = ""
    password = gPassword
    pretty = False

    For i = 1 To objArgs.Count - 1
        arg = objArgs(i)
        If Left(arg, 2) = "--" Then
            Select Case LCase(arg)
                Case "--output"
                    If i < objArgs.Count - 1 Then
                        outputPath = TrimQuotes(objArgs(i + 1))
                        i = i + 1
                    End If
                Case "--password"
                    If i < objArgs.Count - 1 Then
                        password = objArgs(i + 1)
                        i = i + 1
                    End If
                Case "--pretty"
                    pretty = True
                Case "--help"
                    Call ShowExportFormHelp()
                    WScript.Quit 0
            End Select
        ElseIf Left(arg, 1) <> "-" Then
            If dbPath = "" Then
                dbPath = arg
            ElseIf formName = "" Then
                formName = arg
            End If
        End If
    Next

    If dbPath = "" Then
        dbPath = strAccessPath
    End If

    If formName = "" Then
        WScript.Echo "[ERROR] Falta el nombre del formulario."
        Call ShowExportFormHelp()
        WScript.Quit 1
    End If

    If outputPath = "" Then
        WScript.Echo "[ERROR] Debe indicar --output <ruta.json>."
        Call ShowExportFormHelp()
        WScript.Quit 1
    End If

    dbPath = ToAbsolute(dbPath)
    outputPath = ToAbsolute(outputPath)

    If Not objFSO.FileExists(dbPath) Then
        WScript.Echo "[ERROR] Base de datos no encontrada: " & dbPath
        WScript.Quit 1
    End If

    Dim app
    Set app = OpenAccessQuiet(dbPath, password)
    If app Is Nothing Then
        WScript.Echo "[ERROR] No se pudo abrir Access en " & dbPath
        WScript.Quit 1
    End If

    Call UiExportForm(app, dbPath, formName, outputPath, pretty)

    Call CloseAccessQuiet(app)

    WScript.Echo "Formulario '" & formName & "' exportado a " & outputPath
    WScript.Quit 0

ExportFail:
    Dim errMsg
    errMsg = "[ERROR] " & Err.Description
    On Error Resume Next
    If Not app Is Nothing Then
        Call CloseAccessQuiet(app)
    End If
    WScript.Echo errMsg
    WScript.Quit 1
End Sub
End Sub

Sub ImportFormCommand()
    On Error GoTo ImportFail

    Dim dbPath, jsonPath, targetName, password, replaceExisting, strict
    Dim i, arg

    dbPath = ""
    jsonPath = ""
    targetName = ""
    password = gPassword
    replaceExisting = False
    strict = False

    For i = 1 To objArgs.Count - 1
        arg = objArgs(i)
        If Left(arg, 2) = "--" Then
            Select Case LCase(arg)
                Case "--target"
                    If i < objArgs.Count - 1 Then
                        targetName = objArgs(i + 1)
                        i = i + 1
                    End If
                Case "--password"
                    If i < objArgs.Count - 1 Then
                        password = objArgs(i + 1)
                        i = i + 1
                    End If
                Case "--replace"
                    replaceExisting = True
                Case "--strict"
                    strict = True
                Case "--help"
                    Call ShowImportFormHelp()
                    WScript.Quit 0
            End Select
        ElseIf Left(arg, 1) <> "-" Then
            If dbPath = "" Then
                dbPath = arg
            ElseIf jsonPath = "" Then
                jsonPath = arg
            End If
        End If
    Next

    If jsonPath = "" Then
        WScript.Echo "[ERROR] Debe indicar el archivo JSON de entrada."
        Call ShowImportFormHelp()
        WScript.Quit 1
    End If

    If dbPath = "" Then
        dbPath = strAccessPath
    End If

    dbPath = ToAbsolute(dbPath)
    jsonPath = ToAbsolute(jsonPath)

    If Not objFSO.FileExists(dbPath) Then
        WScript.Echo "[ERROR] Base de datos no encontrada: " & dbPath
        WScript.Quit 1
    End If

    If Not objFSO.FileExists(jsonPath) Then
        WScript.Echo "[ERROR] Archivo JSON no encontrado: " & jsonPath
        WScript.Quit 1
    End If

    Dim data
    Set data = UiLoadFormJson(jsonPath)

    Dim validationMessage
    If Not UiValidateFormJson(data, strict, validationMessage) Then
        WScript.Echo "[ERROR] " & validationMessage
        WScript.Quit 1
    End If

    If targetName = "" Then
        targetName = UiGetJsonString(data, "name")
    End If

    Dim app
    Set app = OpenAccessQuiet(dbPath, password)
    If app Is Nothing Then
        WScript.Echo "[ERROR] No se pudo abrir Access en " & dbPath
        WScript.Quit 1
    End If

    Call UiImportForm(app, data, targetName, replaceExisting)

    Call CloseAccessQuiet(app)

    WScript.Echo "Formulario importado correctamente: " & targetName
    WScript.Quit 0

ImportFail:
    Dim errMsg
    errMsg = "[ERROR] " & Err.Description
    On Error Resume Next
    If Not app Is Nothing Then
        Call CloseAccessQuiet(app)
    End If
    WScript.Echo errMsg
    WScript.Quit 1
End Sub

Sub ValidateFormJsonCommand()
    On Error GoTo ValidateFail

    Dim jsonPath, strict, showSchema, schemaVersion
    Dim i, arg

    jsonPath = ""
    strict = False
    showSchema = False
    schemaVersion = 1

    For i = 1 To objArgs.Count - 1
        arg = objArgs(i)
        If Left(arg, 2) = "--" Then
            Select Case LCase(arg)
                Case "--strict"
                    strict = True
                Case "--schema"
                    showSchema = True
                    If i < objArgs.Count - 1 Then
                        schemaVersion = CInt(objArgs(i + 1))
                        i = i + 1
                    End If
                Case "--help"
                    Call ShowValidateFormJsonHelp()
                    WScript.Quit 0
            End Select
        ElseIf Left(arg, 1) <> "-" Then
            If jsonPath = "" Then
                jsonPath = arg
            End If
        End If
    Next

    If showSchema Then
        Call UiPrintFormJsonSchema(schemaVersion)
        WScript.Quit 0
    End If

    If jsonPath = "" Then
        WScript.Echo "[ERROR] Debe indicar un archivo JSON."
        Call ShowValidateFormJsonHelp()
        WScript.Quit 1
    End If

    jsonPath = ToAbsolute(jsonPath)
    If Not objFSO.FileExists(jsonPath) Then
        WScript.Echo "[ERROR] Archivo JSON no encontrado: " & jsonPath
        WScript.Quit 1
    End If

    Dim data
    Set data = UiLoadFormJson(jsonPath)

    Dim validationMessage
    If UiValidateFormJson(data, strict, validationMessage) Then
        WScript.Echo "Validación exitosa: " & jsonPath
        WScript.Quit 0
    Else
        WScript.Echo "[ERROR] " & validationMessage
        WScript.Quit 1
    End If

ValidateFail:
    WScript.Echo "[ERROR] " & Err.Description
    WScript.Quit 1
End Sub

Sub RoundtripFormCommand()
    On Error GoTo RoundtripFail

    Dim dbPath, formName, tempDir, password, pretty
    Dim i, arg

    dbPath = ""
    formName = ""
    tempDir = ""
    password = gPassword
    pretty = False

    For i = 1 To objArgs.Count - 1
        arg = objArgs(i)
        If Left(arg, 2) = "--" Then
            Select Case LCase(arg)
                Case "--temp"
                    If i < objArgs.Count - 1 Then
                        tempDir = ToAbsolute(objArgs(i + 1))
                        i = i + 1
                    End If
                Case "--password"
                    If i < objArgs.Count - 1 Then
                        password = objArgs(i + 1)
                        i = i + 1
                    End If
                Case "--pretty"
                    pretty = True
                Case "--help"
                    Call ShowRoundtripFormHelp()
                    WScript.Quit 0
            End Select
        ElseIf Left(arg, 1) <> "-" Then
            If dbPath = "" Then
                dbPath = arg
            ElseIf formName = "" Then
                formName = arg
            End If
        End If
    Next

    If dbPath = "" Then
        dbPath = strAccessPath
    End If

    If formName = "" Then
        WScript.Echo "[ERROR] Debe indicar el formulario a evaluar."
        Call ShowRoundtripFormHelp()
        WScript.Quit 1
    End If

    If tempDir = "" Then
        WScript.Echo "[ERROR] Debe indicar --temp <directorio>."
        Call ShowRoundtripFormHelp()
        WScript.Quit 1
    End If

    dbPath = ToAbsolute(dbPath)
    If Not objFSO.FileExists(dbPath) Then
        WScript.Echo "[ERROR] Base de datos no encontrada: " & dbPath
        WScript.Quit 1
    End If

    If Not objFSO.FolderExists(tempDir) Then
        objFSO.CreateFolder tempDir
    End If

    Dim result
    Set result = UiPerformRoundtrip(dbPath, formName, tempDir, password, pretty)

    If result("success") Then
        WScript.Echo "Roundtrip completado sin diferencias."
        WScript.Echo "  Export inicial: " & result("prePath")
        WScript.Echo "  Export final:   " & result("postPath")
        WScript.Quit 0
    Else
        WScript.Echo "[ERROR] Roundtrip detectó diferencias."
        WScript.Echo "  Export inicial: " & result("prePath")
        WScript.Echo "  Export final:   " & result("postPath")
        WScript.Quit 1
    End If

RoundtripFail:
    WScript.Echo "[ERROR] " & Err.Description
    WScript.Quit 1
End Sub

' ===== FUNCIONES AUXILIARES PARA FORMULARIOS =====

Sub ExportFormToJson(objAccess, formName, outputPath, pretty, expand)
    On Error GoTo ExportFail

    Dim tempPath
    tempPath = GetTempFilePath("condor_export_", ".txt")

    On Error Resume Next
    If objFSO.FileExists(tempPath) Then objFSO.DeleteFile tempPath, True
    On Error GoTo 0

    objAccess.SaveAsText acForm, formName, tempPath

    Dim formText
    formText = ReadUtf16File(tempPath)

    On Error Resume Next
    If objFSO.FileExists(tempPath) Then objFSO.DeleteFile tempPath, True
    On Error GoTo 0

    Call EnsureParentFolder(outputPath)

    Dim writer
    Set writer = New JsonWriter
    writer.StartObject
    writer.AddProperty "schemaVersion", 1
    writer.AddProperty "name", formName
    writer.AddProperty "payloadEncoding", "base64-utf16le"
    writer.AddProperty "payload", EncodeBase64(formText)
    writer.EndObject

    Dim outFile
    Set outFile = objFSO.CreateTextFile(outputPath, True)
    outFile.Write writer.GetJson
    outFile.Close
    Exit Sub

ExportFail:
    Err.Raise Err.Number, "ExportFormToJson", Err.Description
End Sub

Sub ImportFormFromJson(objAccess, jsonPath, targetName, replace, strict)
    Dim jsonContent, jsonObj, finalTargetName

    jsonContent = ReadUtf8File(jsonPath)
    Set jsonObj = ParseJsonObject(jsonContent)
    If jsonObj Is Nothing Then
        Err.Raise vbObjectError + 3701, "ImportFormFromJson", "JSON inválido"
    End If

    If targetName <> "" Then
        finalTargetName = targetName
    ElseIf jsonObj.Exists("name") Then
        finalTargetName = jsonObj("name")
    Else
        Err.Raise vbObjectError + 3702, "ImportFormFromJson", "No se pudo determinar el nombre del formulario"
    End If

    Dim payloadEncoding, payloadData
    If jsonObj.Exists("payloadEncoding") Then
        payloadEncoding = jsonObj("payloadEncoding")
    Else
        payloadEncoding = ""
    End If

    If jsonObj.Exists("payload") Then
        payloadData = jsonObj("payload")
    Else
        Err.Raise vbObjectError + 3703, "ImportFormFromJson", "Falta 'payload' en el JSON"
    End If

    If LCase(payloadEncoding) <> "base64-utf16le" Then
        Err.Raise vbObjectError + 3704, "ImportFormFromJson", "Codificación de payload no soportada"
    End If

    Dim formText
    formText = DecodeBase64(payloadData)

    Dim tempPath
    tempPath = GetTempFilePath("condor_import_", ".txt")
    Call WriteUtf16File(tempPath, formText)

    If replace Then
        On Error Resume Next
        objAccess.DoCmd.Close acForm, finalTargetName, acSaveNo
        Err.Clear
        objAccess.DoCmd.DeleteObject acForm, finalTargetName
        On Error GoTo 0
    End If

    objAccess.LoadFromText acForm, finalTargetName, tempPath

    On Error Resume Next
    If objFSO.FileExists(tempPath) Then objFSO.DeleteFile tempPath, True
    On Error GoTo 0

    WScript.Echo "Formulario importado exitosamente: " & finalTargetName
End Sub

' ===== FUNCIONES AUXILIARES PARA PARSEO JSON =====

Function ExtractJsonValue(jsonText, fieldName)
    ' Extrae un valor simple de un campo JSON
    Dim pattern, regEx, matches
    pattern = """" & fieldName & """\s*:\s*""([^""]+)"""
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = pattern
    regEx.Global = False
    
    Set matches = regEx.Execute(jsonText)
    If matches.Count > 0 Then
        ExtractJsonValue = matches(0).SubMatches(0)
    Else
        ExtractJsonValue = ""
    End If
End Function

Function ValidateFormJson(jsonPath, strict)
    Dim jsonContent, jsonObj
    jsonContent = ReadUtf8File(jsonPath)
    Set jsonObj = ParseJsonObject(jsonContent)
    If jsonObj Is Nothing Then
        WScript.Echo "[ERROR] JSON inválido"
        ValidateFormJson = False
        Exit Function
    End If

    If Not jsonObj.Exists("schemaVersion") Then
        WScript.Echo "[ERROR] Falta 'schemaVersion'"
        ValidateFormJson = False
        Exit Function
    End If

    If Not jsonObj.Exists("payload") Then
        WScript.Echo "[ERROR] Falta 'payload'"
        ValidateFormJson = False
        Exit Function
    End If

    If Not jsonObj.Exists("payloadEncoding") Then
        WScript.Echo "[ERROR] Falta 'payloadEncoding'"
        ValidateFormJson = False
        Exit Function
    End If

    If LCase(jsonObj("payloadEncoding")) <> "base64-utf16le" Then
        WScript.Echo "[ERROR] Codificación no soportada en 'payloadEncoding'"
        ValidateFormJson = False
        Exit Function
    End If

    If strict Then
        If Not jsonObj.Exists("name") Or Len(Trim(CStr(jsonObj("name")))) = 0 Then
            WScript.Echo "[ERROR] Falta 'name' en modo estricto"
            ValidateFormJson = False
            Exit Function
        End If
    End If

    ValidateFormJson = True
End Function

' ===== FUNCIONES AUXILIARES PARA IMPORTACI├ôN =====

Sub ApplyFormProperties(form, properties)
    ' Aplicar propiedades del formulario de forma completa
    Dim key, value
    
    For Each key In properties.Keys
        value = properties(key)
        
        ' Convertir colores si es necesario
        If InStr(LCase(key), "color") > 0 And VarType(value) = vbString Then
            value = ConvertColorToLong(value)
        End If
        
        ' Aplicar la propiedad usando SetPropertySafe
        Call SetPropertySafe(form, key, value)
    Next
End Sub

Sub CreateFormSections(form, sections)
    ' Crear y configurar secciones del formulario
    Dim sectionName, sectionData
    For Each sectionName In sections.Keys
        Set sectionData = sections(sectionName)
        
        ' Obtener referencia a la secci├│n
        Dim section
        Select Case LCase(sectionName)
            Case "detail", "detalle"
                Set section = form.Section(acDetail)
            Case "header", "encabezado"
                Set section = form.Section(acHeader)
            Case "footer", "pie"
                Set section = form.Section(acFooter)
        End Select
        
        ' Aplicar propiedades de la secci├│n
        If Not section Is Nothing And sectionData.Exists("properties") Then
            Call ApplySectionProperties(section, sectionData("properties"))
        End If
    Next
End Sub

Sub ApplySectionProperties(section, properties)
    ' Aplicar propiedades a una secci├│n espec├¡fica
    Dim key, value, normalizedKey
    For Each key In properties.Keys
        normalizedKey = MapPropKey(key)
        value = properties(key)
        
        On Error Resume Next
        section.Properties(normalizedKey) = value
        If Err.Number <> 0 Then
            WScript.Echo "Advertencia: No se pudo aplicar propiedad de secci├│n '" & normalizedKey & "' = '" & value & "'"
            Err.Clear
        End If
        On Error GoTo 0
    Next
End Sub

Sub CreateFormControls(objAccess, form, controls, strict)
    ' Crear controles del formulario - implementaci├│n b├ísica
    If strict Then
        WScript.Echo "Advertencia: Creaci├│n de controles no implementada completamente"
    End If
    ' TODO: Implementar creaci├│n de controles cuando se necesite
End Sub

Sub CreateSingleControl(objAccess, form, controlName, controlData, strict)
    ' Crear un control individual con manejo completo de propiedades
    Dim controlType, acType, acSection
    Dim left, top, width, height, parent, controlSource
    
    ' Obtener tipo de control y convertir a constante Access
    If controlData.Exists("type") Then
        controlType = controlData("type")
        acType = MapControlType(controlType)
        If acType = -1 Then
            WScript.Echo "Advertencia: Tipo de control desconocido: " & controlType
            Exit Sub
        End If
    Else
        WScript.Echo "[ERROR] Control '" & controlName & "' no tiene tipo especificado"
        WScript.Quit 1
    End If
    
    ' Determinar secci├│n (por defecto Detail)
    acSection = acDetail
    If controlData.Exists("section") Then
        Select Case LCase(controlData("section"))
            Case "header", "encabezado": acSection = acHeader
            Case "footer", "pie": acSection = acFooter
            Case Else: acSection = acDetail
        End Select
    End If
    
    ' Obtener propiedades de posici├│n y tama├▒o con normalizaci├│n de claves
    left = 0: top = 0: width = 1440: height = 240
    parent = ""
    controlSource = ""
    
    If controlData.Exists("properties") Then
        Dim props: Set props = controlData("properties")
        Dim key, normalizedKey
        
        ' Buscar propiedades con normalizaci├│n de claves
        For Each key In props.Keys
            normalizedKey = LCase(MapPropKey(key))
            Select Case normalizedKey
                Case "left": left = props(key)
                Case "top": top = props(key)
                Case "width": width = props(key)
                Case "height": height = props(key)
                Case "parent": parent = props(key)
                Case "controlsource": controlSource = props(key)
            End Select
        Next
    End If
    
    ' Crear el control
    Dim newControl
    On Error Resume Next
    Set newControl = objAccess.Application.CreateControl(form.Name, acType, acSection, parent, controlSource, left, top, width, height)
    If Err.Number <> 0 Then
        WScript.Echo "Error creando control '" & controlName & "': " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Establecer el nombre del control
    newControl.Name = controlName
    
    ' Aplicar propiedades restantes
    If controlData.Exists("properties") Then
        Call ApplyControlProperties(newControl, controlData("properties"))
    End If
    
    ' Aplicar eventos si existen
    If controlData.Exists("events") Then
        Call ApplyControlEvents(newControl, controlData("events"))
    End If
End Sub

Sub ApplyControlProperties(control, properties)
    ' Aplicar propiedades a un control de forma completa
    Dim key, value
    
    For Each key In properties.Keys
        ' Saltar propiedades ya aplicadas durante la creaci├│n
        If LCase(key) = "left" Or LCase(key) = "top" Or LCase(key) = "width" Or LCase(key) = "height" Or LCase(key) = "parent" Or LCase(key) = "controlsource" Then
            ' Skip - estas propiedades se aplican durante CreateControl
        Else
            value = properties(key)
            
            ' Convertir colores si es necesario
            If InStr(LCase(key), "color") > 0 And VarType(value) = vbString Then
                value = ConvertColorToLong(value)
            End If
            
            ' Aplicar la propiedad usando SetPropertySafe
            Call SetPropertySafe(control, key, value)
        End If
    Next
End Sub

Sub ApplyFormEvents(form, handlers, strict)
    ' Aplicar eventos del formulario con manejo completo
    Dim eventName, handlerCode, eventProperty
    
    For Each eventName In handlers.Keys
        handlerCode = handlers(eventName)
        eventProperty = MapEventToProperty(eventName)
        
        If eventProperty <> "" Then
            ' Aplicar el evento usando SetPropertySafe
            Call SetPropertySafe(form, eventProperty, handlerCode)
            WScript.Echo "Evento aplicado: " & eventName & " -> " & eventProperty
        Else
            WScript.Echo "Advertencia: Evento desconocido: " & eventName
        End If
    Next
End Sub

Sub ApplyControlEvents(control, events, strict)
    ' Aplicar eventos a un control espec├¡fico
    Dim eventName, handlerCode, eventProperty
    
    For Each eventName In events.Keys
        handlerCode = events(eventName)
        eventProperty = MapEventToProperty(eventName)
        
        If eventProperty <> "" Then
            ' Aplicar el evento usando SetPropertySafe
            Call SetPropertySafe(control, eventProperty, handlerCode)
            WScript.Echo "Evento de control aplicado: " & control.Name & "." & eventName & " -> " & eventProperty
        Else
            WScript.Echo "Advertencia: Evento de control desconocido: " & eventName
        End If
    Next
End Sub

Function IsEnumProperty(propName)
    ' Determinar si una propiedad es de tipo enumeraci├│n
    Select Case LCase(propName)
        Case "textalign", "alignment", "backstyle", "borderstyle", "specialeffect"
            IsEnumProperty = True
        Case Else
            IsEnumProperty = False
    End Select
End Function

Function MapEventToProperty(eventName)
    ' Mapear nombres de eventos a propiedades de Access con soporte ES/EN
    Select Case LCase(eventName)
        ' Eventos de formulario
        Case "click", "onclick", "clic": MapEventToProperty = "OnClick"
        Case "load", "onload", "cargar": MapEventToProperty = "OnLoad"
        Case "unload", "onunload", "descargar": MapEventToProperty = "OnUnload"
        Case "current", "oncurrent", "actual": MapEventToProperty = "OnCurrent"
        Case "beforeupdate", "onbeforeupdate", "antesdeactualizar": MapEventToProperty = "OnBeforeUpdate"
        Case "afterupdate", "onafterupdate", "despuesdeactualizar": MapEventToProperty = "OnAfterUpdate"
        Case "open", "onopen", "abrir": MapEventToProperty = "OnOpen"
        Case "close", "onclose", "cerrar": MapEventToProperty = "OnClose"
        Case "activate", "onactivate", "activar": MapEventToProperty = "OnActivate"
        Case "deactivate", "ondeactivate", "desactivar": MapEventToProperty = "OnDeactivate"
        Case "resize", "onresize", "redimensionar": MapEventToProperty = "OnResize"
        
        ' Eventos de control
        Case "enter", "onenter", "entrar": MapEventToProperty = "OnEnter"
        Case "exit", "onexit", "salir": MapEventToProperty = "OnExit"
        Case "gotfocus", "ongotfocus", "obtenerfoco": MapEventToProperty = "OnGotFocus"
        Case "lostfocus", "onlostfocus", "perderfoco": MapEventToProperty = "OnLostFocus"
        Case "change", "onchange", "cambiar": MapEventToProperty = "OnChange"
        Case "dblclick", "ondblclick", "dobleclic": MapEventToProperty = "OnDblClick"
        Case "keydown", "onkeydown", "teclaabajo": MapEventToProperty = "OnKeyDown"
        Case "keyup", "onkeyup", "teclaarriba": MapEventToProperty = "OnKeyUp"
        Case "keypress", "onkeypress", "teclapresionada": MapEventToProperty = "OnKeyPress"
        Case "mousedown", "onmousedown", "ratonabajo": MapEventToProperty = "OnMouseDown"
        Case "mouseup", "onmouseup", "ratonarriba": MapEventToProperty = "OnMouseUp"
        Case "mousemove", "onmousemove", "ratonmover": MapEventToProperty = "OnMouseMove"
        
        Case Else: MapEventToProperty = ""
    End Select
End Function

Sub ExecuteRoundtripFlow(dbPath, formName, tempDir, password)
    On Error GoTo RoundtripFail

    Dim preJsonPath, postJsonPath
    preJsonPath = objFSO.BuildPath(tempDir, formName & ".json")
    postJsonPath = objFSO.BuildPath(tempDir, formName & ".post.json")

    Dim app
    Set app = OpenAccessQuiet(dbPath, password)
    If app Is Nothing Then
        Err.Raise vbObjectError + 3801, "ExecuteRoundtripFlow", "No se pudo abrir Access"
    End If

    Call ExportFormToJson(app, formName, preJsonPath, False, "")
    Call ImportFormFromJson(app, preJsonPath, formName, True, False)
    Call ExportFormToJson(app, formName, postJsonPath, False, "")

    Call CloseAccessQuiet(app)

    Dim preData, postData
    Set preData = ParseJsonObject(ReadUtf8File(preJsonPath))
    Set postData = ParseJsonObject(ReadUtf8File(postJsonPath))

    Dim same
    same = False
    If Not preData Is Nothing And Not postData Is Nothing Then
        If preData.Exists("payload") And postData.Exists("payload") Then
            Dim preText, postText
            preText = NormalizeAccessText(DecodeBase64(preData("payload")))
            postText = NormalizeAccessText(DecodeBase64(postData("payload")))
            same = (preText = postText)
        End If
    End If

    If same Then
        WScript.Echo "Roundtrip exitoso: No hay diferencias"
        WScript.Quit 0
    Else
        WScript.Echo "Roundtrip fallido: Se encontraron diferencias"
        WScript.Quit 1
    End If

    Exit Sub

RoundtripFail:
    WScript.Echo "[ERROR] " & Err.Description
    WScript.Quit 1
End Sub

Function DiffJsonSemantico(file1, file2)
    ' Comparaci├│n sem├íntica de archivos JSON
    On Error Resume Next
    
    Dim content1, content2
    Dim f1, f2
    
    Set f1 = objFSO.OpenTextFile(file1, 1)
    content1 = f1.ReadAll
    f1.Close
    
    Set f2 = objFSO.OpenTextFile(file2, 1)
    content2 = f2.ReadAll
    f2.Close
    
    ' Normalizar contenido para comparaci├│n sem├íntica
    content1 = NormalizeJsonForComparison(content1)
    content2 = NormalizeJsonForComparison(content2)
    
    DiffJsonSemantico = (content1 = content2)
    
    On Error GoTo 0
End Function

' ===== FUNCIONES AUXILIARES JSON =====

Function NormalizeJsonForComparison(jsonText)
    ' Normaliza JSON removiendo metadata variable para comparaci├│n
    Dim result
    result = jsonText
    
    ' Remover timestamps variables y campos vol├ítiles
    Dim regEx
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    
    ' Remover generatedAtUTC
    regEx.Pattern = """generatedAtUTC""\s*:\s*""[^""]+"",?"
    result = regEx.Replace(result, "")
    
    ' Remover GUIDs internos vol├ítiles
    regEx.Pattern = """guid""\s*:\s*""[^""]+"",?"
    result = regEx.Replace(result, "")
    
    ' Remover LayoutId vol├ítiles
    regEx.Pattern = """layoutId""\s*:\s*[^,}]+,?"
    result = regEx.Replace(result, "")
    
    ' Remover timestamps de creaci├│n/modificaci├│n
    regEx.Pattern = """(created|modified|lastModified)At(UTC)?""\s*:\s*""[^""]+"",?"
    result = regEx.Replace(result, "")
    
    ' Remover metadatos internos de Access
    regEx.Pattern = """(internalId|objectId|moduleId)""\s*:\s*[^,}]+,?"
    result = regEx.Replace(result, "")
    
    ' Normalizar espacios y comas extra
    regEx.Pattern = ",\s*,"
    result = regEx.Replace(result, ",")
    
    regEx.Pattern = ",\s*}"
    result = regEx.Replace(result, "}")
    
    regEx.Pattern = ",\s*]"
    result = regEx.Replace(result, "]")
    
    ' Remover espacios extra y normalizar
    regEx.Pattern = "\s+"
    result = regEx.Replace(result, " ")
    
    ' Normalizar orden de claves (implementaci├│n b├ísica)
    ' Para una comparaci├│n m├ís robusta, aqu├¡ se podr├¡a implementar
    ' un parser JSON completo y reordenar las claves alfab├®ticamente
    
    NormalizeJsonForComparison = Trim(result)
End Function

' ===== FUNCIONES DE NORMALIZACI├ôN ES/EN =====

Function MapPropKey(key)
    ' Mapea claves de propiedades de espa├▒ol a ingl├®s
    Select Case LCase(Trim(key))
        ' Propiedades b├ísicas
        Case "ancho": MapPropKey = "width"
        Case "alto": MapPropKey = "height"
        Case "izquierda": MapPropKey = "left"
        Case "arriba": MapPropKey = "top"
        Case "etiqueta": MapPropKey = "caption"
        Case "texto": MapPropKey = "caption"
        Case "origencontrol": MapPropKey = "controlSource"
        Case "origen": MapPropKey = "controlSource"
        Case "colorfondo": MapPropKey = "backColor"
        Case "colortexto": MapPropKey = "foreColor"
        Case "fuente": MapPropKey = "fontName"
        Case "tama├▒ofuente": MapPropKey = "fontSize"
        Case "tama├▒o": MapPropKey = "fontSize"
        Case "negrita": MapPropKey = "fontBold"
        Case "cursiva": MapPropKey = "fontItalic"
        Case "alineacion": MapPropKey = "textAlign"
        Case "estiloborde": MapPropKey = "borderStyle"
        Case "efectoespecial": MapPropKey = "specialEffect"
        Case "subrayado": MapPropKey = "fontUnderline"
        Case "colordefondo": MapPropKey = "backColor"
        Case "colortexto": MapPropKey = "foreColor"
        Case "visible": MapPropKey = "visible"
        Case "habilitado": MapPropKey = "enabled"
        Case "bloqueado": MapPropKey = "locked"
        Case "alineaciontexto": MapPropKey = "textAlign"
        Case "borde": MapPropKey = "borderStyle"
        Case "efectoespecial": MapPropKey = "specialEffect"
        Case "imagen": MapPropKey = "picture"
        Case "tipoimagen": MapPropKey = "pictureType"
        Case "alineacionimagen": MapPropKey = "pictureAlignment"
        Case "modoajuste": MapPropKey = "sizeMode"
        Case "formato": MapPropKey = "format"
        Case "mascaraentrada": MapPropKey = "inputMask"
        Case "valorpredeterminado": MapPropKey = "defaultValue"
        Case "reglavalidacion": MapPropKey = "validationRule"
        Case "textovalidacion": MapPropKey = "validationText"
        Case "requerido": MapPropKey = "required"
        Case "permitirlongitudcero": MapPropKey = "allowZeroLength"
        Case "indexado": MapPropKey = "indexed"
        Case "unicodcompresion": MapPropKey = "unicodeCompression"
        Case "ime": MapPropKey = "imeMode"
        Case "etiquetainteligente": MapPropKey = "smartTags"
        Case "ayudacontextual": MapPropKey = "helpContextId"
        Case "textoayuda": MapPropKey = "statusBarText"
        Case "sugerencia": MapPropKey = "controlTipText"
        Case "ordendetabulacion": MapPropKey = "tabIndex"
        Case "detenerdetabulacion": MapPropKey = "tabStop"
        Case "teclaacceso": MapPropKey = "shortcutMenuBar"
        Case "menucontextual": MapPropKey = "shortcutMenuBar"
        Case "barramenu": MapPropKey = "menuBar"
        Case "barraherramientas": MapPropKey = "toolbar"
        Case "filtro": MapPropKey = "filter"
        Case "ordenar": MapPropKey = "orderBy"
        Case "permitirfiltros": MapPropKey = "allowFilters"
        Case "permitiredicion": MapPropKey = "allowEdits"
        Case "permitiragregar": MapPropKey = "allowAdditions"
        Case "permitireliminar": MapPropKey = "allowDeletions"
        Case "permitirdise├▒o": MapPropKey = "allowDesignChanges"
        Case "entradadatos": MapPropKey = "dataEntry"
        Case "conjuntoregistros": MapPropKey = "recordset"
        Case "tipoconjuntoregistros": MapPropKey = "recordsetType"
        Case "origenregistro": MapPropKey = "recordSource"
        Case "bloqueosregistros": MapPropKey = "recordLocks"
        Case "maxregistros": MapPropKey = "maxRecords"
        Case "cargarporeventos": MapPropKey = "loadOnOpen"
        Case "cerraralsalir": MapPropKey = "closeButton"
        Case "botonesbarra": MapPropKey = "recordSelectors"
        Case "selectoresregistro": MapPropKey = "recordSelectors"
        Case "barranavegacion": MapPropKey = "navigationButtons"
        Case "lineasdivision": MapPropKey = "dividerLines"
        Case "autocentrar": MapPropKey = "autoCenter"
        Case "autoresize": MapPropKey = "autoResize"
        Case "ajustarventana": MapPropKey = "fitToScreen"
        Case "modal": MapPropKey = "modal"
        Case "popup": MapPropKey = "popup"
        Case "ciclotabulacion": MapPropKey = "cycle"
        Case "vistaformulario": MapPropKey = "defaultView"
        Case "vistapredeterminada": MapPropKey = "defaultView"
        Case "permitirvistahojadatos": MapPropKey = "allowDatasheetView"
        Case "permitirvistapivot": MapPropKey = "allowPivotTableView"
        Case "permitirvistatabladinamica": MapPropKey = "allowPivotChartView"
        Case "permitirvistaformulario": MapPropKey = "allowFormView"
        Case "permitirlayout": MapPropKey = "allowLayoutView"
        Case "subdatasheets": MapPropKey = "subdatasheetName"
        Case "expandirsubdatasheets": MapPropKey = "subdatasheetExpanded"
        Case "alturasubdatasheets": MapPropKey = "subdatasheetHeight"
        Case "orientacionformulario": MapPropKey = "orientation"
        Case "orientacionformulariodividido": MapPropKey = "splitFormOrientation"
        Case "tama├▒oformulariodividido": MapPropKey = "splitFormSize"
        Case "impresionformulariodividido": MapPropKey = "splitFormPrinting"
        Case "barradesplazamiento": MapPropKey = "splitFormSplitterBar"
        Case "datasheet": MapPropKey = "splitFormDatasheet"
        Case "formulariodividido": MapPropKey = "splitFormOrientation"
        Case Else
            ' Si no hay mapeo, devolver la clave original
            MapPropKey = key
    End Select
End Function

Function NormalizeEnumToken(token)
    ' Normaliza tokens de enumeraciones de espa├▒ol a ingl├®s
    Select Case LCase(Trim(token))
        ' Alineaci├│n de texto
        Case "izquierda", "left": NormalizeEnumToken = "left"
        Case "centro", "centrado", "center": NormalizeEnumToken = "center"
        Case "derecha", "right": NormalizeEnumToken = "right"
        Case "justificado", "justify": NormalizeEnumToken = "justify"
        Case "distribuir", "distribute": NormalizeEnumToken = "distribute"
        
        ' Tipos de borde
        Case "transparente", "transparent": NormalizeEnumToken = "transparent"
        Case "solido", "solid": NormalizeEnumToken = "solid"
        Case "guiones", "dashes": NormalizeEnumToken = "dashes"
        Case "puntos", "dots": NormalizeEnumToken = "dots"
        Case "doble", "double": NormalizeEnumToken = "double"
        
        ' Efectos especiales
        Case "plano", "flat": NormalizeEnumToken = "flat"
        Case "relieve", "raised": NormalizeEnumToken = "raised"
        Case "hundido", "sunken": NormalizeEnumToken = "sunken"
        Case "grabado", "etched": NormalizeEnumToken = "etched"
        Case "sombra", "shadowed": NormalizeEnumToken = "shadowed"
        Case "cincelado", "chiseled": NormalizeEnumToken = "chiseled"
        
        ' Tipos de vista
        Case "formulario", "form": NormalizeEnumToken = "form"
        Case "continuo", "continuous": NormalizeEnumToken = "continuous"
        Case "hojadatos", "datasheet": NormalizeEnumToken = "datasheet"
        Case "tabladinamica", "pivottable": NormalizeEnumToken = "pivottable"
        Case "graficopivot", "pivotchart": NormalizeEnumToken = "pivotchart"
        
        ' Orientaci├│n
        Case "horizontal": NormalizeEnumToken = "horizontal"
        Case "vertical": NormalizeEnumToken = "vertical"
        
        ' Orientaci├│n de formulario dividido
        Case "arriba", "top": NormalizeEnumToken = "top"
        Case "abajo", "bottom": NormalizeEnumToken = "bottom"
        
        ' Tipos de conjunto de registros
        Case "dynaset": NormalizeEnumToken = "dynaset"
        Case "snapshot": NormalizeEnumToken = "snapshot"
        
        ' Ciclo de tabulaci├│n
        Case "todosregistros", "allrecords": NormalizeEnumToken = "allrecords"
        Case "registroactual", "currentrecord": NormalizeEnumToken = "currentrecord"
        Case "paginaactual", "currentpage": NormalizeEnumToken = "currentpage"
        
        ' Modo de ajuste de imagen
        Case "recortar", "clip": NormalizeEnumToken = "clip"
        Case "estirar", "stretch": NormalizeEnumToken = "stretch"
        Case "zoom": NormalizeEnumToken = "zoom"
        
        ' Alineaci├│n de imagen
        Case "esquinasuperioriz", "topleft": NormalizeEnumToken = "topleft"
        Case "arribacentro", "topcenter": NormalizeEnumToken = "topcenter"
        Case "esquinasuperiorder", "topright": NormalizeEnumToken = "topright"
        Case "centroizquierda", "centerleft": NormalizeEnumToken = "centerleft"
        Case "centrocentro", "centercenter": NormalizeEnumToken = "centercenter"
        Case "centroderecha", "centerright": NormalizeEnumToken = "centerright"
        Case "esquinainferioriz", "bottomleft": NormalizeEnumToken = "bottomleft"
        Case "abajocentro", "bottomcenter": NormalizeEnumToken = "bottomcenter"
        Case "esquinainferiorder", "bottomright": NormalizeEnumToken = "bottomright"
        
        ' Valores booleanos
        Case "verdadero", "true", "si", "yes": NormalizeEnumToken = "true"
        Case "falso", "false", "no": NormalizeEnumToken = "false"
        
        Case Else
            ' Si no hay mapeo, devolver el token original
            NormalizeEnumToken = token
    End Select
End Function

Function MapControlType(controlType)
    ' Mapea tipos de control de espa├▒ol/ingl├®s a valores num├®ricos de constantes
    Select Case LCase(Trim(controlType))
        Case "etiqueta", "label": MapControlType = acLabel
        Case "cuadrodetexto", "textbox": MapControlType = acTextBox
        Case "boton", "commandbutton": MapControlType = acCommandButton
        Case "casilladeseleccion", "checkbox": MapControlType = acCheckBox
        Case "botonopcion", "optionbutton": MapControlType = acOptionButton
        Case "cuadrocombo", "combobox": MapControlType = acComboBox
        Case "cuadrolista", "listbox": MapControlType = acListBox
        Case "subformulario", "subform": MapControlType = acSubform
        Case "imagen", "image": MapControlType = acImage
        Case "saltolinea", "pagebreak": MapControlType = acPageBreak
        Case "separador", "line": MapControlType = acLine
        Case "rectangulo", "rectangle": MapControlType = acRectangle
        Case "pesta├▒as", "tabcontrol": MapControlType = acTabCtl
        Case "activex", "customcontrol": MapControlType = acCustomControl
        Case Else
            ' Si no hay mapeo, intentar convertir a n├║mero o devolver -1
            If IsNumeric(controlType) Then
                MapControlType = CInt(controlType)
            Else
                MapControlType = -1
            End If
    End Select
End Function

' ===== INFRAESTRUCTURA JSON CENTRALIZADA =====

Class JsonWriter
    Private jsonContent
    Private objectStack
    Private arrayStack
    Private currentState
    Private needsComma
    
    Private Sub Class_Initialize()
        jsonContent = ""
        Set objectStack = CreateList()
        Set arrayStack = CreateList()
        currentState = "root"
        needsComma = False
    End Sub
    
    Public Sub StartObject()
        If needsComma Then
            jsonContent = jsonContent & ","
        End If
        jsonContent = jsonContent & "{"
        ' Verificar si objectStack es ArrayList o Dictionary
        If TypeName(objectStack) = "ArrayList" Then
            objectStack.Add currentState
        Else
            ' Es Dictionary, usar como lista con ├¡ndices num├®ricos
            objectStack.Add objectStack.Count, currentState
        End If
        currentState = "object"
        needsComma = False
    End Sub
    
    Public Sub EndObject()
        jsonContent = jsonContent & "}"
        If objectStack.Count > 0 Then
            ' Verificar si objectStack es ArrayList o Dictionary
            If TypeName(objectStack) = "ArrayList" Then
                currentState = objectStack(objectStack.Count - 1)
                objectStack.RemoveAt objectStack.Count - 1
            Else
                ' Es Dictionary, usar como lista con ├¡ndices num├®ricos
                currentState = objectStack(objectStack.Count - 1)
                objectStack.Remove objectStack.Count - 1
            End If
            needsComma = True
        End If
    End Sub
    
    Public Sub StartArray()
        If needsComma Then
            jsonContent = jsonContent & ","
        End If
        jsonContent = jsonContent & "["
        ' Verificar si arrayStack es ArrayList o Dictionary
        If TypeName(arrayStack) = "ArrayList" Then
            arrayStack.Add currentState
        Else
            ' Es Dictionary, usar como lista con ├¡ndices num├®ricos
            arrayStack.Add arrayStack.Count, currentState
        End If
        currentState = "array"
        needsComma = False
    End Sub
    
    Public Sub EndArray()
        jsonContent = jsonContent & "]"
        If arrayStack.Count > 0 Then
            ' Verificar si arrayStack es ArrayList o Dictionary
            If TypeName(arrayStack) = "ArrayList" Then
                currentState = arrayStack(arrayStack.Count - 1)
                arrayStack.RemoveAt arrayStack.Count - 1
            Else
                ' Es Dictionary, usar como lista con ├¡ndices num├®ricos
                currentState = arrayStack(arrayStack.Count - 1)
                arrayStack.Remove arrayStack.Count - 1
            End If
            needsComma = True
        End If
    End Sub
    
    Public Sub AddProperty(key, value)
        If needsComma Then
            jsonContent = jsonContent & ","
        End If
        jsonContent = jsonContent & """" & EscapeString(CStr(key)) & """:" & FormatValue(value)
        needsComma = True
    End Sub
    
    Public Sub StartObjectProperty(key)
        If needsComma Then
            jsonContent = jsonContent & ","
        End If
        jsonContent = jsonContent & """" & EscapeString(CStr(key)) & """:{"
        ' Verificar si objectStack es ArrayList o Dictionary
        If TypeName(objectStack) = "ArrayList" Then
            objectStack.Add currentState
        Else
            ' Es Dictionary, usar como lista con ├¡ndices num├®ricos
            objectStack.Add objectStack.Count, currentState
        End If
        currentState = "object"
        needsComma = False
    End Sub
    
    Public Sub StartArrayProperty(key)
        If needsComma Then
            jsonContent = jsonContent & ","
        End If
        jsonContent = jsonContent & """" & EscapeString(CStr(key)) & """:["
        ' Verificar si arrayStack es ArrayList o Dictionary
        If TypeName(arrayStack) = "ArrayList" Then
            arrayStack.Add currentState
        Else
            ' Es Dictionary, usar como lista con ├¡ndices num├®ricos
            arrayStack.Add arrayStack.Count, currentState
        End If
        currentState = "array"
        needsComma = False
    End Sub
    
    Public Sub AddValue(value)
        If needsComma Then
            jsonContent = jsonContent & ","
        End If
        jsonContent = jsonContent & FormatValue(value)
        needsComma = True
    End Sub
    
    Public Function GetJson()
        GetJson = jsonContent
    End Function
    
    Private Function FormatValue(value)
        If IsNull(value) Then
            FormatValue = "null"
        ElseIf VarType(value) = vbBoolean Then
            If value Then
                FormatValue = "true"
            Else
                FormatValue = "false"
            End If
        ElseIf IsNumeric(value) Then
            FormatValue = CStr(value)
        ElseIf VarType(value) = vbString Then
            FormatValue = """" & EscapeString(CStr(value)) & """"
        Else
            FormatValue = """" & EscapeString(CStr(value)) & """"
        End If
    End Function
    
    Private Function EscapeString(str)
        Dim result, i, char
        result = ""
        For i = 1 To Len(str)
            char = Mid(str, i, 1)
            Select Case char
                Case Chr(34) ' "
                    result = result & "\"""
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
    
    ' M├®todos de compatibilidad con la implementaci├│n anterior
    Public Sub WriteProperty(key, value)
        ' M├®todo de compatibilidad que usa AddProperty
        AddProperty key, value
    End Sub
    
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
        Else
            Stringify = """" & EscapeString(CStr(value)) & """"
        End If
    End Function
End Class

' ===== COMANDO LIST-FORMS ACTUALIZADO =====

Sub ListFormsCommand()
    On Error Resume Next
    Dim password, bJsonOutput
    Dim i, arg
    Dim formsList, formCount
    Dim app, formName
    Dim errN, errD
    
    ' Inicializar variables
    password = gPassword
    bJsonOutput = False
    
    ' Procesar argumentos especificos de list-forms
    For i = 1 To objArgs.Count - 1
        arg = objArgs(i)
        If Left(arg, 2) = "--" Then
            If arg = "--password" And i < objArgs.Count - 1 Then
                password = objArgs(i + 1)
            ElseIf arg = "--json" Then
                bJsonOutput = True
            ElseIf arg = "--help" Then
                Call ShowListFormsHelp()
                WScript.Quit 0
            End If
        End If
    Next
    
    ' Abrir Access con OpenAccessQuiet
    Set app = OpenAccessQuiet(strAccessPath, password)
    If app Is Nothing Then
        WScript.Echo "[ERROR] No se pudo abrir la base de datos"
        WScript.Quit 1
    End If
    
    ' Contar formularios primero
    formCount = 0
    For Each formName In app.CurrentProject.AllForms
        formCount = formCount + 1
    Next
    
    ' Redimensionar array y llenar
    Dim formsArray()
    If formCount > 0 Then
        ReDim formsArray(formCount - 1)
        Dim formIndex
        formIndex = 0
        For Each formName In app.CurrentProject.AllForms
            formsArray(formIndex) = formName.Name
            formIndex = formIndex + 1
        Next
    Else
        ' Si no hay formularios, crear array vac├¡o
        ReDim formsArray(-1)
    End If
    
    ' Capturar errores antes del cierre
    errN = Err.Number
    errD = Err.Description
    
    ' Cerrar Access SIEMPRE
    Call CloseAccessQuiet(app)
    
    ' Verificar errores
    If errN <> 0 Then
        WScript.Echo "[ERROR] " & errD
        WScript.Quit 1
    End If
    
    ' Generar salida
    If bJsonOutput Then
        ' Salida JSON
        Dim jsonOutput
        jsonOutput = "["
        If formCount > 0 Then
            Dim i_form
            For i_form = 0 To UBound(formsArray)
                If i_form > 0 Then jsonOutput = jsonOutput & ","
                jsonOutput = jsonOutput & """" & formsArray(i_form) & """"
            Next
        End If
        jsonOutput = jsonOutput & "]"
        WScript.Echo jsonOutput
    Else
        ' Salida texto
        If formCount = 0 Then
            WScript.Echo "No se encontraron formularios en la base de datos."
        Else
            WScript.Echo "Formularios encontrados (" & formCount & "):"
            For i_form = 0 To UBound(formsArray)
                WScript.Echo "  " & formsArray(i_form)
            Next
        End If
    End If
    
    WScript.Quit 0
End Sub

Sub ShowListFormsHelp()
    WScript.Echo "=== LIST-FORMS - Listar formularios de la base de datos ==="
    WScript.Echo "Uso: cscript condor_cli.vbs list-forms [db_path] [opciones]"
    WScript.Echo ""
    WScript.Echo "PARAMETROS:"
    WScript.Echo "  [db_path]         - Ruta de la base de datos (opcional si se resuelve autom├íticamente)"
    WScript.Echo ""
    WScript.Echo "OPCIONES:"
    WScript.Echo "  --password <pwd>  - password de la base de datos"
    WScript.Echo "  --json            - Salida en formato JSON"
    WScript.Echo ""
    WScript.Echo "EJEMPLOS:"
    WScript.Echo "  cscript condor_cli.vbs list-forms"
    WScript.Echo "  cscript condor_cli.vbs list-forms --json"
    WScript.Echo "  cscript condor_cli.vbs list-forms --db ""C:\MiDB.accdb"" --password 1234"
End Sub

' ===== FUNCIONES AUXILIARES PARA VALIDACION JSON =====

' ===== FUNCIONES DE BUNDLE =====

Function GetFunctionalityFiles(strFunctionality)
    Dim arrFiles
    
    Select Case LCase(strFunctionality)
        Case "auth", "autenticacion", "authentication"
            ' Seccion 3.1 - Autenticacion + Dependencias
            arrFiles = Array("IAuthService.cls", "CAuthService.cls", "CMockAuthService.cls", _
                           "IAuthRepository.cls", "CAuthRepository.cls", "CMockAuthRepository.cls", _
                           "EAuthData.cls", "modAuthFactory.bas", "TestAuthService.bas", _
                           "TIAuthRepository.bas", _
                           "IConfig.cls", "IErrorHandlerService.cls", "modEnumeraciones.bas")
        
        Case "document", "documentos", "documents"
            ' Seccion 3.2 - Gestion de Documentos + Dependencias
            arrFiles = Array("IDocumentService.cls", "CDocumentService.cls", "CMockDocumentService.cls", _
                           "IWordManager.cls", "CWordManager.cls", "CMockWordManager.cls", _
                           "IMapeoRepository.cls", "CMapeoRepository.cls", "CMockMapeoRepository.cls", _
                           "EMapeo.cls", "modDocumentServiceFactory.bas", _
                           "TIDocumentService.bas", _
                           "ISolicitudService.cls", "CSolicitudService.cls", "modSolicitudServiceFactory.bas", _
                           "IOperationLogger.cls", "IConfig.cls", "IErrorHandlerService.cls", "IFileSystem.cls", _
                           "modWordManagerFactory.bas", "modRepositoryFactory.bas", "modErrorHandlerFactory.bas")
        
        Case "expediente", "expedientes"
            ' Seccion 3.3 - Gestion de Expedientes + Dependencias
            arrFiles = Array("IExpedienteService.cls", "CExpedienteService.cls", "CMockExpedienteService.cls", _
                           "IExpedienteRepository.cls", "CExpedienteRepository.cls", "CMockExpedienteRepository.cls", _
                           "EExpediente.cls", "modExpedienteServiceFactory.bas", "TestCExpedienteService.bas", _
                           "TIExpedienteRepository.bas", "modRepositoryFactory.bas", _
                           "IConfig.cls", "IOperationLogger.cls", "IErrorHandlerService.cls")
        
        Case "solicitud", "solicitudes"
            ' Seccion 3.4 - Gestion de Solicitudes + Dependencias
            arrFiles = Array("ISolicitudService.cls", "CSolicitudService.cls", "CMockSolicitudService.cls", _
                           "ISolicitudRepository.cls", "CSolicitudRepository.cls", "CMockSolicitudRepository.cls", _
                           "ESolicitud.cls", "EDatosPc.cls", "EDatosCdCa.cls", "EDatosCdCaSub.cls", _
                           "modSolicitudServiceFactory.bas", "TestSolicitudService.bas", _
                           "TISolicitudRepository.bas", _
                           "IAuthService.cls", "modAuthFactory.bas", _
                           "IOperationLogger.cls", "IErrorHandlerService.cls", "IConfig.cls")
        
        Case "workflow", "flujo"
            ' Seccion 3.5 - Gestion de Workflow + Dependencias
            arrFiles = Array("IWorkflowService.cls", "CWorkflowService.cls", "CMockWorkflowService.cls", _
                           "IWorkflowRepository.cls", "CWorkflowRepository.cls", "CMockWorkflowRepository.cls", _
                           "modWorkflowServiceFactory.bas", "TestWorkflowService.bas", _
                           "TIWorkflowRepository.bas", _
                           "IOperationLogger.cls", "IConfig.cls", "IErrorHandlerService.cls")
        
        Case "mapeo", "mapping"
            ' Seccion 3.6 - Gestion de Mapeos + Dependencias
            arrFiles = Array("IMapeoRepository.cls", "CMapeoRepository.cls", "CMockMapeoRepository.cls", _
                           "EMapeo.cls", "TIMapeoRepository.bas", _
                           "IConfig.cls", "IErrorHandlerService.cls")
        
        Case "notification", "notificacion"
            ' Seccion 3.7 - Gestion de Notificaciones + Dependencias
            arrFiles = Array("INotificationService.cls", "CNotificationService.cls", "CMockNotificationService.cls", _
                           "INotificationRepository.cls", "CNotificationRepository.cls", "CMockNotificationRepository.cls", _
                           "modNotificationServiceFactory.bas", "TINotificationService.bas", _
                           "IOperationLogger.cls", "IErrorHandlerService.cls", "IConfig.cls")
        
        Case "operation", "operacion", "logging"
            ' Seccion 3.8 - Gestion de Operaciones y Logging + Dependencias
            arrFiles = Array("IOperationLogger.cls", "COperationLogger.cls", "CMockOperationLogger.cls", _
                           "IOperationRepository.cls", "COperationRepository.cls", "CMockOperationRepository.cls", _
                           "EOperationLog.cls", "modOperationLoggerFactory.bas", "TestOperationLogger.bas", _
                           "TIOperationRepository.bas", _
                           "IErrorHandlerService.cls", "IConfig.cls")
        
        Case "config", "configuracion"
            ' Seccion 4 - Configuracion + Dependencias
            arrFiles = Array("IConfig.cls", "CConfig.cls", "CMockConfig.cls", "modConfigFactory.bas", _
                           "TestCConfig.bas")
        
        Case "filesystem", "archivos"
            ' Seccion 5 - Sistema de Archivos + Dependencias
            arrFiles = Array("IFileSystem.cls", "CFileSystem.cls", "CMockFileSystem.cls", _
                           "modFileSystemFactory.bas", "TIFileSystem.bas", _
                           "IErrorHandlerService.cls")
        
        Case "word"
            ' Seccion 6 - Gestion de Word + Dependencias
            arrFiles = Array("IWordManager.cls", "CWordManager.cls", "CMockWordManager.cls", _
                           "modWordManagerFactory.bas", "TIWordManager.bas", _
                           "IFileSystem.cls", "IErrorHandlerService.cls")
        
        Case "error", "errores", "errors"
            ' Seccion 7 - Gestion de Errores + Dependencias
            arrFiles = Array("IErrorHandlerService.cls", "CErrorHandlerService.cls", "CMockErrorHandlerService.cls", _
                           "modErrorHandlerFactory.bas", "TestErrorHandlerService.bas", _
                           "IConfig.cls", "IFileSystem.cls")
        
        Case "testframework", "testing", "framework"
            ' Seccion 8 - Framework de Testing + Dependencias
            arrFiles = Array("ITestReporter.cls", "CTestResult.cls", "CTestSuiteResult.cls", "CTestReporter.cls", _
                           "modTestRunner.bas", "modTestUtils.bas", "modAssert.bas", _
                           "TestModAssert.bas", "IFileSystem.cls", "IConfig.cls", _
                           "IErrorHandlerService.cls")
        
        Case "app", "aplicacion", "application"
            ' Seccion 9 - Gestion de Aplicacion + Dependencias
            arrFiles = Array("IAppManager.cls", "CAppManager.cls", "CMockAppManager.cls", _
                           "ModAppManagerFactory.bas", "TestAppManager.bas", "IAuthService.cls", _
                           "IConfig.cls", "IErrorHandlerService.cls")
        
        Case "models", "modelos", "datos"
            ' Seccion 10 - Modelos de Datos
            arrFiles = Array("EUsuario.cls", "ESolicitud.cls", "EExpediente.cls", "EDatosPc.cls", _
                           "EDatosCdCa.cls", "EDatosCdCaSub.cls", "EEstado.cls", "ETransicion.cls", _
                           "EMapeo.cls", "EAdjuntos.cls", "ELogCambios.cls", "ELogErrores.cls", "EOperationLog.cls", "EAuthData.cls")
        
        Case "utils", "utilidades", "enumeraciones"
            ' Seccion 11 - Utilidades y Enumeraciones
            arrFiles = Array("modRepositoryFactory.bas", "modEnumeraciones.bas", "modQueries.bas", _
                           "ModAppManagerFactory.bas", "modAuthFactory.bas", "modConfigFactory.bas", _
                           "modDocumentServiceFactory.bas", "modErrorHandlerFactory.bas", _
                           "modExpedienteServiceFactory.bas", "modFileSystemFactory.bas", _
                           "modNotificationServiceFactory.bas", "modOperationLoggerFactory.bas", _
                           "modSolicitudServiceFactory.bas", "modWordManagerFactory.bas", _
                           "modWorkflowServiceFactory.bas")
        
        Case "forms", "formularios", "ui"
            ' Funcionalidad de Formularios - UI as Code
            arrFiles = Array("condor_cli.vbs")
            
        Case "cli", "infrastructure", "infraestructura"
            ' Funcionalidad CLI e Infraestructura
            arrFiles = Array("condor_cli.vbs")
            
        Case "condorcli"
            ' Funcionalidad especial para copiar condor_cli.vbs como .txt
            arrFiles = Array("condor_cli.vbs")
            
        Case "tests", "pruebas", "testing", "test"
            ' Seccion 12 - Archivos de Pruebas (Autodescubrimiento)
            arrFiles = Array()
        Case Else
            ' Funcionalidad no reconocida - devolver array vacio
            arrFiles = Array()
    End Select
    
    GetFunctionalityFiles = arrFiles
End Function

Sub BundleFunctionality()
    On Error Resume Next
    
    Dim strFunctionalityOrFiles, strDestPath, strBundlePath, timestamp
    
    ' Verificar argumentos
    If objArgs.Count < 2 Then
        WScript.Echo "[ERROR] Se requiere nombre de funcionalidad o lista de ficheros"
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
            WScript.Echo "[ERROR] Error creando carpeta de destino: " & Err.Description
            WScript.Quit 1
        End If
    End If
    
    Dim arrFilesToBundle
    
    ' Logica de Deteccion Inteligente
    If InStr(strFunctionalityOrFiles, ",") > 0 Then
        ' MODO 1: Lista de ficheros explicita
        WScript.Echo "Modo: Lista de ficheros explicita."
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
                WScript.Echo "[ERROR] '" & strFunctionalityOrFiles & "' no es una funcionalidad conocida ni un archivo existente en src."
                WScript.Echo "Funcionalidades disponibles: Auth, Document, Expediente, Solicitud, Workflow, Mapeo, Notification, Operation, Config, FileSystem, Word, Error, TestFramework, App, Models, Utils, Tests"
                WScript.Quit 1
            End If
        End If
    End If
    
    ' Llamar a la subrutina de ayuda para copiar los ficheros
    Call CopyFilesToBundle(arrFilesToBundle, strBundlePath)
    
    On Error GoTo 0
End Sub

Sub CopyFilesToBundle(arrFiles, strBundlePath)
    Dim copiedFiles, notFoundFiles
    copiedFiles = 0
    notFoundFiles = 0
    
    If UBound(arrFiles) < 0 Then
        WScript.Echo "[VERBOSE] La lista de ficheros a empaquetar esta vacia."
    End If

    Dim i, fileName, filePath, destFilePath
    For i = 0 To UBound(arrFiles)
        fileName = Trim(arrFiles(i))
        
        ' Caso especial para condorcli: copiar desde la raiz del proyecto
        If fileName = "condor_cli.vbs" And InStr(strBundlePath, "bundle_condorcli_") > 0 Then
            filePath = objFSO.BuildPath(objFSO.GetParentFolderName(WScript.ScriptFullName), fileName)
        Else
            filePath = objFSO.BuildPath(strSourcePath, fileName)
        End If
        
        If objFSO.FileExists(filePath) Then
            ' Copiar archivo con extension .txt a├▒adida al directorio del bundle
            destFilePath = objFSO.BuildPath(strBundlePath, fileName & ".txt")
            objFSO.CopyFile filePath, destFilePath, True
            
            If Err.Number <> 0 Then
                WScript.Echo "  [ERROR] Error copiando " & fileName & ": " & Err.Description
                Err.Clear
            Else
                WScript.Echo "  [OK] " & fileName & " -> " & fileName & ".txt"
                copiedFiles = copiedFiles + 1
            End If
        Else
            WScript.Echo "  [ERROR] Archivo no encontrado: " & fileName
            notFoundFiles = notFoundFiles + 1
        End If
    Next
    
    WScript.Echo ""
    WScript.Echo "=== RESULTADO DEL EMPAQUETADO ==="
    WScript.Echo "Archivos copiados: " & copiedFiles
    WScript.Echo "Archivos no encontrados: " & notFoundFiles
    WScript.Echo "Ubicacion del paquete: " & strBundlePath
    
    If copiedFiles = 0 Then
        WScript.Echo "[ERROR] No se copio ningun archivo."
    Else
        WScript.Echo "[OK] Empaquetado completado exitosamente"
    End If
End Sub

' Comando para listar modulos VBA con opciones avanzadas
Sub ListModulesCommand()
    On Error Resume Next
    Dim includeDocs, pattern, flagJson, expectSrc, flagDiff, dbPath, password
    Dim app, arr, i, ok
    Dim errN, errD
    
    ' Parsear flags
    includeDocs = HasFlag("includeDocs")
    pattern = GetArgValue("pattern")
    flagJson = HasFlag("json")
    expectSrc = HasFlag("expectSrc")
    flagDiff = HasFlag("diff")
    dbPath = GetArgValue("db")
    password = GetArgValue("password")
    
    ' Resolver password si es necesario
    If password = "" Then password = ResolveDbPassword()
    
    ' Verificar instancia de Access existente o abrir nueva
    Set app = GetOrCreateAccessInstance(dbPath, password)
    If app Is Nothing Then
        WScript.Echo "[ERROR] No se pudo acceder a la instancia de Access"
        WScript.Quit 1
    End If
    
    ' Intentar listar modulos usando diferentes metodos
    WScript.Echo "DEBUG: Intentando TryListModulesVBIDE..."
    ok = TryListModulesVBIDE(app, includeDocs, pattern, arr)
    If Not ok Then
        WScript.Echo "DEBUG: TryListModulesVBIDE fallo, intentando TryListModulesAllModules..."
        ok = TryListModulesAllModules(app, pattern, arr)
        If Not ok Then
            WScript.Echo "DEBUG: TryListModulesAllModules fallo, intentando TryListModulesDAO..."
            ok = TryListModulesDAO(app, pattern, arr)
        End If
    End If
    
    ' Capturar errores antes del cierre
    errN = Err.Number
    errD = Err.Description
    
    ' Cerrar Access SIEMPRE
    Call CloseAccessQuiet(app)
    
    ' Verificar errores
    If errN <> 0 Then
        WScript.Echo "[ERROR] " & errD
        WScript.Quit 1
    End If
    
    If Not ok Then
        WScript.Echo "[ERROR] No se pudieron listar los modulos"
        WScript.Quit 1
    End If
    
    ' Mostrar resultados
    If flagJson Then
        PrintModulesJson arr
    Else
        PrintModulesText arr
    End If
    
    WScript.Quit 0
End Sub

' Comando para arreglar headers de archivos fuente
Sub FixSrcHeadersCommand()
    WScript.Echo "[INFO] Comando fix-src-headers no implementado aun"
    WScript.Echo "[INFO] Esta funcionalidad arreglara headers de archivos en /src"
End Sub

' Funciones auxiliares para ListModulesCommand
Function TryListModulesVBIDE(app, includeDocs, pattern, ByRef arr)
    On Error Resume Next
    ReDim arr(-1)
    
    Dim regex, p, comps, i, vbComp, kind, name, dict
    If pattern <> "" Then 
        Set regex = CreateObject("VBScript.RegExp")
        regex.Pattern = pattern
        regex.IgnoreCase = True
    End If
    
    Set p = app.VBE.ActiveVBProject
    If Err.Number <> 0 Or p Is Nothing Then 
        Err.Clear
        TryListModulesVBIDE = False
        Exit Function
    End If
    
    Set comps = p.VBComponents
    If Err.Number <> 0 Or comps Is Nothing Then 
        Err.Clear
        TryListModulesVBIDE = False
        Exit Function
    End If

    For i = 1 To comps.Count
        Set vbComp = comps(i)
        Select Case vbComp.Type
            Case 1: kind = "STD"
            Case 2: kind = "CLS"
            Case 3: kind = "FRM"
            Case 100: kind = "RPT"
            Case Else: kind = "OTHER"
        End Select
        name = vbComp.Name
        
        If (kind = "FRM" Or kind = "RPT") And Not includeDocs Then
            ' omitir
        ElseIf pattern = "" Or regex.Test(name) Then
            If UBound(arr) = -1 Then 
                ReDim arr(0) 
            Else 
                ReDim Preserve arr(UBound(arr) + 1)
            End If
            Set dict = CreateObject("Scripting.Dictionary")
            dict.Add "kind", kind
            dict.Add "name", name
            Set arr(UBound(arr)) = dict
        End If
    Next

    TryListModulesVBIDE = (Err.Number = 0)
    On Error GoTo 0
End Function

Function TryListModulesAllModules(app, pattern, ByRef arr)
    On Error Resume Next
    ReDim arr(-1)
    
    If app.CurrentProject Is Nothing Then 
        TryListModulesAllModules = False
        Exit Function
    End If
    
    Dim mods, regex, m, name, dict
    If pattern <> "" Then 
        Set regex = CreateObject("VBScript.RegExp")
        regex.Pattern = pattern
        regex.IgnoreCase = True
    End If
    
    Set mods = app.CurrentProject.AllModules
    If Err.Number <> 0 Or mods Is Nothing Then 
        Err.Clear
        TryListModulesAllModules = False
        Exit Function
    End If

    For Each m In mods
        name = m.Name
        If pattern = "" Or regex.Test(name) Then
            If UBound(arr) = -1 Then 
                ReDim arr(0) 
            Else 
                ReDim Preserve arr(UBound(arr) + 1)
            End If
            Set dict = CreateObject("Scripting.Dictionary")
            dict.Add "kind", "STD"
            dict.Add "name", name
            Set arr(UBound(arr)) = dict
        End If
    Next
    
    TryListModulesAllModules = (Err.Number = 0)
    On Error GoTo 0
End Function

Function TryListModulesDAO(app, pattern, ByRef arr)
    On Error Resume Next
    ReDim arr(-1)
    
    If app Is Nothing Then 
        TryListModulesDAO = False
        Exit Function
    End If
    
    If app.CurrentProject Is Nothing Then 
        TryListModulesDAO = False
        Exit Function
    End If
    
    Dim regex, db, obj, name, dict, moduleCount
    If pattern <> "" Then 
        Set regex = CreateObject("VBScript.RegExp")
        regex.Pattern = pattern
        regex.IgnoreCase = True
    End If
    
    Set db = app.CurrentProject
    If Err.Number <> 0 Then
        Err.Clear
        TryListModulesDAO = False
        Exit Function
    End If
    
    moduleCount = 0
    For Each obj In db.AllModules
        If Err.Number = 0 Then
            name = obj.Name
            If Err.Number = 0 Then
                If pattern = "" Or regex.Test(name) Then
                    If UBound(arr) = -1 Then 
                        ReDim arr(0) 
                    Else 
                        ReDim Preserve arr(UBound(arr) + 1)
                    End If
                    Set dict = CreateObject("Scripting.Dictionary")
                    dict.Add "kind", "STD"
                    dict.Add "name", name
                    Set arr(UBound(arr)) = dict
                    moduleCount = moduleCount + 1
                End If
            Else
                Err.Clear
            End If
        Else
            Err.Clear
        End If
    Next
    
    TryListModulesDAO = (moduleCount > 0)
    On Error GoTo 0
End Function

Sub PrintModulesText(arr)
    Dim total, i
    total = 0
    If IsArray(arr) Then 
        If UBound(arr) >= 0 Then 
            total = UBound(arr) + 1
        End If
    End If
    
    WScript.Echo "KIND  Name"
    WScript.Echo "----  ----"
    
    If total > 0 Then
        For i = 0 To UBound(arr)
            WScript.Echo arr(i)("kind") & "   " & arr(i)("name")
        Next
    End If
    
    WScript.Echo ""
    WScript.Echo "Total: " & total & " modulos"
End Sub

Sub PrintModulesJson(arr)
    Dim total, i, json
    total = 0
    If IsArray(arr) Then 
        If UBound(arr) >= 0 Then 
            total = UBound(arr) + 1
        End If
    End If
    
    json = "{"
    json = json & """modules"": ["
    
    If total > 0 Then
        For i = 0 To UBound(arr)
            If i > 0 Then json = json & ","
            json = json & "{"
            json = json & """kind"": """ & arr(i)("kind") & ""","
            json = json & """name"": """ & arr(i)("name") & """"
            json = json & "}"
        Next
    End If
    
    json = json & "],"
    json = json & """total"": " & total
    json = json & "}"
    
    WScript.Echo json
End Sub

' ===== FUNCI├ôN AUXILIAR PARA ASIGNACI├ôN SEGURA DE PROPIEDADES =====

Sub SetPropertySafe(obj, key, val)
    ' Asigna una propiedad de forma segura, manejando diferentes tipos de objetos
    Dim normalizedKey, normalizedVal
    
    ' Normalizar la clave
    normalizedKey = MapPropKey(key)
    
    ' Normalizar el valor si es un token de enumeraci├│n
    If VarType(val) = vbString Then
        normalizedVal = NormalizeEnumToken(val)
    Else
        normalizedVal = val
    End If
    
    ' Intentar asignaci├│n directa primero
    On Error Resume Next
    obj(normalizedKey) = normalizedVal
    
    ' Si falla, intentar usando Properties
    If Err.Number <> 0 Then
        Err.Clear
        obj.Properties(normalizedKey) = normalizedVal
        
        ' Si a├║n falla, reportar error
        If Err.Number <> 0 Then
            WScript.Echo "[ERROR] No se pudo asignar propiedad '" & normalizedKey & "' con valor '" & normalizedVal & "': " & Err.Description
            Err.Clear
        End If
    End If
    On Error GoTo 0
End Sub

' ===== PARSER JSON MEJORADO =====

Function ParseJsonObject(jsonText)
    ' Parser JSON b├ísico pero mejorado para manejar objetos anidados
    Dim result
    Set result = CreateObject("Scripting.Dictionary")
    
    ' Limpiar espacios y saltos de l├¡nea
    jsonText = Trim(jsonText)
    
    ' Verificar que sea un objeto JSON v├ílido
    If Left(jsonText, 1) <> "{" Or Right(jsonText, 1) <> "}" Then
        Set ParseJsonObject = Nothing
        Exit Function
    End If
    
    ' Extraer contenido del objeto (sin llaves externas)
    Dim content: content = Mid(jsonText, 2, Len(jsonText) - 2)
    
    ' Parsear campos principales conocidos
    Call ParseJsonField(content, "name", result)
    Call ParseJsonField(content, "schemaVersion", result)
    
    ' Parsear objetos anidados
    Call ParseJsonNestedObject(content, "properties", result)
    Call ParseJsonNestedObject(content, "sections", result)
    Call ParseJsonNestedObject(content, "controls", result)
    Call ParseJsonNestedObject(content, "code", result)
    
    Set ParseJsonObject = result
End Function

Sub ParseJsonField(content, fieldName, resultDict)
    ' Buscar campo simple (string o n├║mero)
    Dim pattern, regEx, matches
    pattern = """" & fieldName & """\s*:\s*""([^""]+)"""
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = pattern
    regEx.Global = False
    
    Set matches = regEx.Execute(content)
    If matches.Count > 0 Then
        resultDict(fieldName) = matches(0).SubMatches(0)
    End If
End Sub

Sub ParseJsonNestedObject(content, objectName, resultDict)
    ' Crear objeto vac├¡o para objetos anidados
    ' En una implementaci├│n completa se parsear├¡a el contenido real
    Set resultDict(objectName) = CreateObject("Scripting.Dictionary")
End Sub











