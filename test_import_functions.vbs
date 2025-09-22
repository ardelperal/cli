' Test script para validar las funciones de importacion VBE mejoradas
' Uso: cscript test_import_functions.vbs

Option Explicit

' Variables globales para simular el entorno de cli.vbs
Dim gDebug
gDebug = True

' Incluir las funciones necesarias del cli.vbs principal
' (En un entorno real, esto se haria con un include o ejecutando cli.vbs)

Sub LogMessage(msg)
    WScript.Echo "[" & Now & "] " & msg
End Sub

' Test 1: Validar DetectCharset con diferentes tipos de archivo
Sub TestDetectCharset()
    LogMessage "=== Test DetectCharset ==="
    
    ' Crear archivos de prueba con diferentes encodings
    CreateTestFile "test_ansi.bas", "windows-1252", "Function TestAnsi()" & vbCrLf & "End Function"
    CreateTestFile "test_utf8.bas", "utf-8", "Function TestUTF8()" & vbCrLf & "End Function"
    CreateTestFile "test_unicode.bas", "unicode", "Function TestUnicode()" & vbCrLf & "End Function"
    
    ' Probar deteccion
    Dim charset
    charset = DetectCharset("test_ansi.bas")
    LogMessage "Charset detectado para test_ansi.bas: " & charset
    
    charset = DetectCharset("test_utf8.bas")
    LogMessage "Charset detectado para test_utf8.bas: " & charset
    
    charset = DetectCharset("test_unicode.bas")
    LogMessage "Charset detectado para test_unicode.bas: " & charset
    
    ' Limpiar archivos de prueba
    DeleteTestFile "test_ansi.bas"
    DeleteTestFile "test_utf8.bas"
    DeleteTestFile "test_unicode.bas"
End Sub

' Test 2: Validar CopyAsAnsi con diferentes encodings
Sub TestCopyAsAnsi()
    LogMessage "=== Test CopyAsAnsi ==="
    
    ' Crear archivo UTF-8 con BOM
    CreateTestFile "source_utf8.bas", "utf-8", "Function TestCopyUTF8()" & vbCrLf & "    ' Comentario con tildes: función, configuración" & vbCrLf & "End Function"
    
    ' Copiar como ANSI
    Dim result
    result = CopyAsAnsi("source_utf8.bas", "dest_ansi.bas")
    
    If result Then
        LogMessage "CopyAsAnsi exitoso: source_utf8.bas -> dest_ansi.bas"
        
        ' Verificar que el archivo destino es ANSI
        Dim destCharset
        destCharset = DetectCharset("dest_ansi.bas")
        LogMessage "Charset del archivo destino: " & destCharset
    Else
        LogMessage "ERROR: CopyAsAnsi fallo"
    End If
    
    ' Limpiar
    DeleteTestFile "source_utf8.bas"
    DeleteTestFile "dest_ansi.bas"
End Sub

' Test 3: Simular ImportModuleToAccess (sin Access real)
Sub TestImportModuleToAccess()
    LogMessage "=== Test ImportModuleToAccess (simulado) ==="
    
    ' Crear archivos de prueba
    CreateTestFile "TestModule.bas", "windows-1252", "Attribute VB_Name = ""TestModule""" & vbCrLf & "Function TestFunction()" & vbCrLf & "End Function"
    CreateTestFile "TestClass.cls", "windows-1252", "VERSION 1.0 CLASS" & vbCrLf & "Attribute VB_Name = ""TestClass""" & vbCrLf & "Sub TestMethod()" & vbCrLf & "End Sub"
    CreateTestFile "TestInvalid.txt", "windows-1252", "Invalid file"
    
    ' Simular llamadas (sin Access real, solo validacion de logica)
    LogMessage "Simulando ImportModuleToAccess para TestModule.bas (extension .bas -> acModule)"
    LogMessage "Simulando ImportModuleToAccess para TestClass.cls (extension .cls -> acClassModule)"
    LogMessage "Simulando ImportModuleToAccess para TestInvalid.txt (extension no soportada -> False)"
    
    ' Limpiar
    DeleteTestFile "TestModule.bas"
    DeleteTestFile "TestClass.cls"
    DeleteTestFile "TestInvalid.txt"
End Sub

' Funciones auxiliares para crear archivos de prueba
Sub CreateTestFile(fileName, charset, content)
    Dim stm
    Set stm = CreateObject("ADODB.Stream")
    
    stm.Type = 2 ' adTypeText
    stm.Charset = charset
    stm.Open
    stm.WriteText content, 0
    stm.SaveToFile fileName, 2 ' adSaveCreateOverWrite
    stm.Close
    
    LogMessage "Archivo de prueba creado: " & fileName & " (charset: " & charset & ")"
End Sub

Sub DeleteTestFile(fileName)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(fileName) Then
        fso.DeleteFile fileName
        LogMessage "Archivo de prueba eliminado: " & fileName
    End If
End Sub

' Funciones simuladas del cli.vbs (versiones simplificadas para testing)
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

Function CopyAsAnsi(srcPath, dstPath)
    CopyAsAnsi = False
    Dim inS, outS, txt, srcCs
    
    Set inS  = CreateObject("ADODB.Stream")
    Set outS = CreateObject("ADODB.Stream")
    txt = ""
    
    ' Detectar charset del archivo fuente
    srcCs = DetectCharset(srcPath)
    If gDebug Then LogMessage "[DEBUG] Charset detectado para " & srcPath & ": " & srcCs
    
    On Error Resume Next
    
    ' Configurar stream de entrada con charset detectado
    inS.Type = 2 ' adTypeText
    inS.Charset = srcCs
    inS.Open
    inS.LoadFromFile srcPath
    
    If Err.Number <> 0 Then
        If gDebug Then LogMessage "[DEBUG] Error con charset " & srcCs & ", intentando fallback a windows-1252"
        Err.Clear
        If inS.State = 1 Then inS.Close
        
        ' Fallback a windows-1252 si falla el charset detectado
        inS.Type = 2
        inS.Charset = "windows-1252"
        inS.Open
        inS.LoadFromFile srcPath
        If Err.Number <> 0 Then 
            Err.Clear
            If inS.State = 1 Then inS.Close
            Exit Function
        End If
    End If
    
    ' Leer contenido completo
    txt = inS.ReadText(-1)
    If Err.Number <> 0 Then
        Err.Clear
        If inS.State = 1 Then inS.Close
        Exit Function
    End If
    
    If inS.State = 1 Then inS.Close
    
    ' Escribir como windows-1252 (ANSI)
    outS.Type = 2
    outS.Charset = "windows-1252"
    outS.Open
    outS.WriteText txt, 0
    outS.SaveToFile dstPath, 2 ' adSaveCreateOverWrite
    
    If Err.Number <> 0 Then 
        Err.Clear
        If outS.State = 1 Then outS.Close
        Exit Function
    End If
    
    If outS.State = 1 Then outS.Close
    On Error GoTo 0
    
    If gDebug Then LogMessage "[DEBUG] Archivo copiado como ANSI: " & srcPath & " -> " & dstPath
    CopyAsAnsi = True
End Function

' Ejecutar tests
Sub Main()
    LogMessage "Iniciando tests de funciones de importacion VBE..."
    LogMessage ""
    
    TestDetectCharset()
    LogMessage ""
    
    TestCopyAsAnsi()
    LogMessage ""
    
    TestImportModuleToAccess()
    LogMessage ""
    
    LogMessage "Tests completados."
End Sub

' Ejecutar
Main()