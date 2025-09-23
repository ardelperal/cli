Attribute VB_Name = "modTestUtils"
Option Compare Database
Option Explicit

    Public Function GetProjectPath() As String
        ' Detecta "\front\" o "\back\" en la ruta del proyecto y devuelve la raíz del repo.
        Dim full As String: full = CurrentProject.FullName
        Dim pFront As Long: pFront = InStrRev(full, "\front\", -1, vbTextCompare)
        Dim pBack  As Long: pBack = InStrRev(full, "\back\", -1, vbTextCompare)
        Dim cutPos As Long

        If pFront > 0 Then
            cutPos = pFront - 1
        ElseIf pBack > 0 Then
            cutPos = pBack - 1
        Else
            ' Fallback seguro: carpeta del .accdb
            GetProjectPath = CurrentProject.Path
            Exit Function
        End If

        GetProjectPath = Left$(full, cutPos)
    End Function

    Public Function GetWorkspacePath() As String
        ' El workspace de pruebas vive en FRONT
        GetWorkspacePath = JoinPath(GetProjectPath(), "front\test_env\workspace\")
    End Function
    
    Public Function JoinPath(ByVal basePath As String, ByVal relativePath As String) As String
        Dim b As String: b = Trim$(basePath)
        Dim r As String: r = Trim$(relativePath)
        If Len(b) = 0 Then JoinPath = r: Exit Function
        If Len(r) = 0 Then JoinPath = b: Exit Function
        If Right$(b, 1) = "\" Then b = Left$(b, Len(b) - 1)
        If Left$(r, 1) = "\" Then r = Mid$(r, 2)
        JoinPath = b & "\" & r
    End Function

    Public Sub ProvisionTestDatabases()
        On Error GoTo ErrorHandler
        Debug.Print "--- INICIO DE APROVISIONAMIENTO CENTRALIZADO ---"
        
        Dim config As IConfig: Set config = modTestContext.GetTestConfig()
        Dim fs As IFileSystem: Set fs = modFileSystemFactory.CreateFileSystem(config)
        
        fs.CreateFolder GetWorkspacePath()
        
        Dim dbsToProvision As Variant
        dbsToProvision = Array("CONDOR", "LANZADERA", "EXPEDIENTES", "CORREOS")
        
        Dim i As Integer, sourceKey As String, destKey As String, sourcePath As String, destPath As String
        
        For i = LBound(dbsToProvision) To UBound(dbsToProvision)
            sourceKey = "DEV_" & dbsToProvision(i) & "_DATA_PATH"
            destKey = dbsToProvision(i) & "_DATA_PATH"
            sourcePath = config.GetValue(sourceKey)
            destPath = config.GetValue(destKey)
    
            Debug.Print "Aprovisionando: Origen=[" & sourcePath & "], Destino=[" & destPath & "]"
            
            ' BLINDAJE: Lanzar un error claro si la clave de destino no se encuentra en modTestContext
            If Len(destPath) = 0 Then
                Err.Raise vbObjectError + 5002, "Provision", "La clave de destino '" & destKey & "' no fue encontrada en modTestContext.bas"
            End If
            
            If Not fs.FileExists(sourcePath) Then
                Err.Raise vbObjectError + 5001, "Provision", "Origen no encontrado: " & sourcePath
            End If
            
            If fs.FileExists(destPath) Then fs.DeleteFile destPath
            
            fs.CopyFile sourcePath, destPath
            
            If Not fs.FileExists(destPath) Then
                Err.Raise vbObjectError + 5003, "Provision", "Copia fallida a: " & destPath
            Else
                Debug.Print "   -> Copia exitosa."
            End If
        Next i

        Debug.Print "--- FIN DE APROVISIONAMIENTO ---"
        Exit Sub
ErrorHandler:
        Debug.Print "--- FALLO CRÍTICO EN APROVISIONAMIENTO: " & Err.Description & " ---"
        Err.Raise Err.Number, "modTestUtils.ProvisionTestDatabases", Err.Description
    End Sub
    
    Public Sub CleanupWorkspace()
        On Error Resume Next
        Dim fs As IFileSystem: Set fs = modFileSystemFactory.CreateFileSystem()
        Dim workspacePath As String: workspacePath = GetWorkspacePath()
        If fs.FolderExists(workspacePath) Then
            fs.DeleteFolderRecursive workspacePath
            fs.CreateFolder workspacePath
        End If
    End Sub

    Public Sub PrintTestConfigStatus()
        On Error GoTo ErrorHandler
        
        Debug.Print "--- INICIO: AUDITORÍA DE CONFIGURACIÓN DE PRUEBAS ---"
        
        Dim config As IConfig
        Set config = modTestContext.GetTestConfig()
        
        If config Is Nothing Then
            Debug.Print "ERROR: No se pudo obtener la configuración de pruebas."
            Exit Sub
        End If
        
        Dim fs As IFileSystem
        Set fs = modFileSystemFactory.CreateFileSystem(config)
        
        ' Para acceder a las claves, necesitamos castear a la implementación concreta
        Dim mockConfig As CMockConfig
        Set mockConfig = config
        
        Dim key As Variant
        Dim value As String
        Dim status As String
        
        For Each key In mockConfig.GetAllKeys() ' Asumimos que CMockConfig expondrá este método
            value = config.GetValue(key)
            status = ""
            
            ' Heurística para detectar si el valor es una ruta de fichero o directorio
            If InStr(value, "\") > 0 And (InStr(value, ".accdb") > 0 Or InStr(value, ".log") > 0 Or Right(value, 1) = "\") Then
                If Right(value, 1) = "\" Then ' Es un directorio
                    If fs.FolderExists(value) Then
                        status = "[? DIRECTORIO OK]"
                    Else
                        status = "[X DIRECTORIO NO ENCONTRADO]"
                    End If
                Else ' Es un fichero
                    If fs.FileExists(value) Then
                        status = "[? FICHERO OK]"
                    Else
                        status = "[X FICHERO NO ENCONTRADO]"
                    End If
                End If
            End If
            
            Debug.Print Left(key & Space(35), 35) & "| " & value & " " & status
        Next key

        Debug.Print "--- FIN: AUDITORÍA DE CONFIGURACIÓN DE PRUEBAS ---"
        
        Set config = Nothing
        Set fs = Nothing
        Set mockConfig = Nothing
        Exit Sub

ErrorHandler:
        Debug.Print "ERROR CRÍTICO durante la auditoría de configuración: " & Err.Description
    End Sub


    'HOLA'
