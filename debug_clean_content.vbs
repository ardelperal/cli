' Script para debuggear el contenido que genera CleanVBAFile
Option Explicit

Dim objFSO, moduleName, sourceFile
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Obtener el nombre del módulo a testear (por defecto CAppManager)
If WScript.Arguments.Count > 0 Then
    moduleName = WScript.Arguments(0)
Else
    moduleName = "CAppManager"
End If

' Buscar el archivo fuente
sourceFile = "C:\Proyectos\cli\src\" & moduleName & ".cls"
If Not objFSO.FileExists(sourceFile) Then
    sourceFile = "C:\Proyectos\cli\src\" & moduleName & ".bas"
    If Not objFSO.FileExists(sourceFile) Then
        WScript.Echo "❌ ERROR: No se encuentra el archivo fuente: " & sourceFile
        WScript.Quit 1
    End If
End If

WScript.Echo "=== DEBUG CLEANVBAFILE ==="
WScript.Echo "Archivo fuente: " & sourceFile

' Leer contenido original
Dim objStream, originalContent
Set objStream = CreateObject("ADODB.Stream")
objStream.Type = 2 ' adTypeText
objStream.Charset = "UTF-8"
objStream.Open
objStream.LoadFromFile sourceFile
originalContent = objStream.ReadText
objStream.Close

WScript.Echo "Contenido original: " & Len(originalContent) & " caracteres"
WScript.Echo "Primeras 10 líneas del archivo original:"
Dim arrLines, i
arrLines = Split(originalContent, vbLf)
For i = 0 To UBound(arrLines)
    If i >= 10 Then Exit For
    WScript.Echo "   " & (i+1) & ": " & arrLines(i)
Next

' Simular CleanVBAFile
Dim cleanedContent, fileExtension
fileExtension = LCase(objFSO.GetExtensionName(sourceFile))
cleanedContent = CleanVBAFile(sourceFile, fileExtension)

WScript.Echo ""
WScript.Echo "Contenido limpio: " & Len(cleanedContent) & " caracteres"
WScript.Echo "Contenido completo limpio:"
WScript.Echo "[" & cleanedContent & "]"

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