Attribute VB_Name = "modHealthCheck"
Option Compare Database
Option Explicit

Public Function GenerateHealthReport() As String
    On Error GoTo ErrorHandler
    
    Dim report As String
    report = "--- INFORME DE SALUD DEL SISTEMA CONDOR (DINÁMICO) ---" & vbCrLf
    report = report & "Fecha: " & Now() & vbCrLf
    report = report & String(70, "-") & vbCrLf & vbCrLf
    
    Dim config As IConfig
    Set config = modConfigFactory.CreateConfigService()
    If config Is Nothing Then
        GenerateHealthReport = "ERROR CRÍTICO: No se pudo cargar el servicio de configuración."
        Exit Function
    End If
    
    Dim fs As IFileSystem
    Set fs = modFileSystemFactory.CreateFileSystem()
    
    ' --- Verificación Dinámica de TODAS las Claves de Configuración ---
    report = report & "1. AUDITORÍA DE CLAVES DE CONFIGURACIÓN:" & vbCrLf
    
    Dim configImpl As CConfig
    Set configImpl = config
    
    Dim key As Variant, value As String, status As String
    
    For Each key In configImpl.GetAllKeys()
        ' DOBLE BLINDAJE: Asegurarse de que la clave no sea nula o vacía antes de procesar
        If Not IsNull(key) And Len(Trim(key & "")) > 0 Then
            value = CStr(config.GetValue(key) & "")
            
            If InStr(key, "PATH") > 0 Then
                If Right(value, 1) = "\" Then
                    If fs.FolderExists(value) Then status = "[? OK]" Else status = "[X ERROR: Directorio no encontrado]"
                Else
                    If fs.FileExists(value) Then status = "[? OK]" Else status = "[X ERROR: Fichero no encontrado]"
                End If
            Else
                status = "[INFO]" ' No es una ruta, solo se informa del valor
            End If
            report = report & "  " & Left(key & Space(35), 35) & ": " & status & " (" & value & ")" & vbCrLf
        End If
    Next key
    
    ' --- Verificación Específica de Plantillas Word ---
    report = report & vbCrLf & "2. VERIFICACIÓN DE PLANTILLAS WORD:" & vbCrLf
    
    Dim templatesPath As String
    templatesPath = CStr(config.GetValue("TEMPLATES_PATH") & "")
    
    If Len(templatesPath) = 0 Or Not fs.FolderExists(templatesPath) Then
        report = report & "  [X ERROR: No se puede verificar plantillas porque TEMPLATES_PATH no es válido.]" & vbCrLf
    Else
        ' Leer los nombres de las plantillas dinámicamente desde la configuración
        Dim templatesToVerifyKeys As Variant
        templatesToVerifyKeys = Array("TEMPLATE_NAME_PC", "TEMPLATE_NAME_CDCA", "TEMPLATE_NAME_CDCASUB")
        
        Dim keyName As Variant
        Dim templateName As String
        Dim templateFullPath As String
        
        For Each keyName In templatesToVerifyKeys
            templateName = CStr(config.GetValue(CStr(keyName)) & "")
            
            If Len(templateName) = 0 Then
                status = "[X ERROR: Clave '" & keyName & "' no configurada]"
                report = report & "  " & Left(CStr(keyName) & Space(55), 55) & ": " & status & vbCrLf
            Else
                templateFullPath = modTestUtils.JoinPath(templatesPath, templateName)
                If fs.FileExists(templateFullPath) Then
                    status = "[? OK]"
                Else
                    status = "[X ERROR: No encontrada]"
                End If
                report = report & "  " & Left(templateName & Space(55), 55) & ": " & status & " (" & templateFullPath & ")" & vbCrLf
            End If
        Next keyName
    End If
    
    report = report & vbCrLf & String(70, "-") & vbCrLf & "Diagnóstico completado." & vbCrLf
    GenerateHealthReport = report
    Exit Function
    
ErrorHandler:
    GenerateHealthReport = "ERROR INESPERADO durante la verificación: " & Err.Description
End Function
