' Script para debuggear ImportModuleWithAnsiEncoding
Option Explicit

Dim objAccess, objFSO, moduleName, sourceFile
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Obtener el nombre del módulo a testear
If WScript.Arguments.Count > 0 Then
    moduleName = WScript.Arguments(0)
Else
    moduleName = "CAppManager"
End If

WScript.Echo "=== DEBUG IMPORT MODULE ==="
WScript.Echo "Módulo: " & moduleName

' Abrir Access
On Error Resume Next
Set objAccess = CreateObject("Access.Application")
If Err.Number <> 0 Then
    WScript.Echo "❌ Error abriendo Access: " & Err.Description
    WScript.Quit 1
End If
On Error GoTo 0

objAccess.Visible = False
objAccess.OpenCurrentDatabase "C:\Proyectos\cli\condor.accdb", False, "dpddpd"

' Buscar el archivo fuente
sourceFile = "C:\Proyectos\cli\src\" & moduleName & ".cls"
If Not objFSO.FileExists(sourceFile) Then
    sourceFile = "C:\Proyectos\cli\src\" & moduleName & ".bas"
    If Not objFSO.FileExists(sourceFile) Then
        WScript.Echo "❌ ERROR: No se encuentra el archivo fuente: " & sourceFile
        objAccess.Quit
        WScript.Quit 1
    End If
End If

Dim fileExtension
fileExtension = LCase(objFSO.GetExtensionName(sourceFile))

' Generar contenido limpio
Dim cleanedContent
cleanedContent = CleanVBAFile(sourceFile, fileExtension)

WScript.Echo "Contenido limpio generado: " & Len(cleanedContent) & " caracteres"
WScript.Echo "Primeras 5 líneas del contenido limpio:"
Dim arrLines, i
arrLines = Split(cleanedContent, vbCrLf)
For i = 0 To UBound(arrLines)
    If i >= 5 Then Exit For
    WScript.Echo "   " & (i+1) & ": " & arrLines(i)
Next

' Verificar si el módulo existe antes de importar
Dim vbProject, vbComponent, moduleExists
Set vbProject = objAccess.VBE.ActiveVBProject
moduleExists = False

For Each vbComponent In vbProject.VBComponents
    If vbComponent.Name = moduleName Then
        moduleExists = True
        WScript.Echo "Módulo " & moduleName & " ya existe, eliminándolo..."
        vbProject.VBComponents.Remove vbComponent
        Exit For
    End If
Next

' Crear nuevo componente
Dim newComponent, componentType
If LCase(fileExtension) = "cls" Then
    componentType = 2  ' vbext_ct_ClassModule
Else
    componentType = 1  ' vbext_ct_StdModule
End If

WScript.Echo "Creando nuevo componente tipo " & componentType & " para: " & moduleName

On Error Resume Next
Set newComponent = vbProject.VBComponents.Add(componentType)
If Err.Number <> 0 Then
    WScript.Echo "❌ Error creando componente: " & Err.Description
    objAccess.Quit
    WScript.Quit 1
End If
On Error GoTo 0

newComponent.Name = moduleName
WScript.Echo "Componente creado con nombre: " & newComponent.Name

' Limpiar el código del componente recién creado
WScript.Echo "Líneas iniciales en el componente: " & newComponent.CodeModule.CountOfLines
If newComponent.CodeModule.CountOfLines > 0 Then
    newComponent.CodeModule.DeleteLines 1, newComponent.CodeModule.CountOfLines
    WScript.Echo "Componente limpiado, líneas restantes: " & newComponent.CodeModule.CountOfLines
End If

' Insertar contenido limpio
WScript.Echo "Insertando contenido de " & Len(cleanedContent) & " caracteres..."
On Error Resume Next
newComponent.CodeModule.AddFromString cleanedContent
If Err.Number <> 0 Then
    WScript.Echo "❌ Error insertando contenido: " & Err.Description
    objAccess.Quit
    WScript.Quit 1
End If
On Error GoTo 0

WScript.Echo "Contenido insertado. Líneas finales: " & newComponent.CodeModule.CountOfLines

' FORZAR GUARDADO DEL PROYECTO VBA
WScript.Echo "💾 Guardando proyecto VBA..."
On Error Resume Next
objAccess.DoCmd.Save acModule, moduleName
If Err.Number <> 0 Then
    WScript.Echo "⚠️ Error al guardar módulo: " & Err.Description
    Err.Clear
End If

' Guardar el proyecto completo
objAccess.DoCmd.RunCommand 2040  ' acCmdSaveRecord
If Err.Number <> 0 Then
    WScript.Echo "⚠️ Error al guardar proyecto: " & Err.Description
    Err.Clear
End If

' Guardar la base de datos
objAccess.DoCmd.Save
If Err.Number <> 0 Then
    WScript.Echo "⚠️ Error al guardar base de datos: " & Err.Description
    Err.Clear
End If
On Error GoTo 0

WScript.Echo "✓ Guardado completado"

' Verificar el contenido final
Dim finalContent
finalContent = ""
For i = 1 To newComponent.CodeModule.CountOfLines
    finalContent = finalContent & newComponent.CodeModule.Lines(i, 1) & vbCrLf
Next

WScript.Echo "Contenido final en VBE: " & Len(finalContent) & " caracteres"
WScript.Echo "Primeras 5 líneas del contenido final:"
arrLines = Split(finalContent, vbCrLf)
For i = 0 To UBound(arrLines)
    If i >= 5 Then Exit For
    WScript.Echo "   " & (i+1) & ": " & arrLines(i)
Next

objAccess.Quit
WScript.Echo "=== FIN DEBUG ==="

' Función CleanVBAFile copiada del cli.vbs
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
    Dim hasOptionCompareDatabase, hasOptionExplicit
    hasOptionCompareDatabase = False
    hasOptionExplicit = False
    
    ' Iterar sobre el array de líneas original
    For i = 0 To UBound(arrLines)
        strLine = arrLines(i)
        
        ' Detectar si ya existe Option Compare Database u Option Explicit
        If Trim(strLine) = "Option Compare Database" Then
            hasOptionCompareDatabase = True
        End If
        If Trim(strLine) = "Option Explicit" Then
            hasOptionExplicit = True
        End If
        
        ' Aplicar las reglas para descartar contenido no deseado
        ' Una línea se descarta si cumple cualquiera de estas condiciones:
        ' CORRECCION CRITICA: Filtrar TODAS las líneas que empiecen con 'Attribute'
        ' y todos los metadatos de archivos .cls
        ' PRESERVAR: Option Compare Database y Option Explicit son esenciales
        If Not (Left(Trim(strLine), 9) = "Attribute" Or _
                Left(Trim(strLine), 17) = "VERSION 1.0 CLASS" Or _
                Trim(strLine) = "BEGIN" Or _
                Left(Trim(strLine), 8) = "MultiUse" Or _
                Trim(strLine) = "END") Then
            
            ' Si no cumple ninguna condición, es código VBA válido
            ' Se añade al cleanedContent seguida de un salto de línea
            cleanedContent = cleanedContent & strLine & vbCrLf
        End If
    Next
    
    ' HOTFIX: Si cleanedContent queda vacío, devolver el contenido original
    ' pero al menos preservar Option Compare Database si no existe
    If Trim(cleanedContent) = "" Then
        cleanedContent = strContent
        WScript.Echo "WARN: CleanVBAFile devolvió vacío para " & filePath & ", usando contenido original"
    Else
        ' Asegurar que Option Compare Database esté presente si no existe
        If Not hasOptionCompareDatabase And fileExtension = ".bas" Then
            cleanedContent = "Option Compare Database" & vbCrLf & cleanedContent
        End If
    End If
    
    ' La función devuelve cleanedContent directamente
    ' No añade ninguna cabecera Option manualmente
    CleanVBAFile = cleanedContent
End Function