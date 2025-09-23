VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IFileSystem

' =====================================================
' CLASE: CFileSystem
' DESCRIPCI?N: Implementaci?n real del sistema de ficheros
' AUTOR: Sistema CONDOR
' FECHA: 2024
' VERSION: 1.0 - Wrapper del FileSystemObject
' =====================================================

' Wrapper alrededor del FileSystemObject real
' Permite abstraer las operaciones de sistema de ficheros
' para facilitar las pruebas unitarias con mocks

Private fso As Object ' Scripting.FileSystemObject
Private m_Config As IConfig

' ============================================================================
' CONSTRUCTOR Y DESTRUCTOR
' ============================================================================

Private Sub Class_Terminate()
    Set fso = Nothing
    Set m_Config = Nothing
End Sub

' ============================================================================
' INICIALIZACI?N CON INYECCI?N DE DEPENDENCIAS
' ============================================================================

Public Sub Initialize(ByVal config As IConfig)
    Set m_Config = config
    Set fso = CreateObject("Scripting.FileSystemObject")
End Sub

' ============================================================================
' IMPLEMENTACI?N DE LA INTERFAZ IFileSystem
' ============================================================================

' Implementa la apertura de archivos de texto
Private Function IFileSystem_OpenTextFile(ByVal Path As String, ByVal Mode As Long, ByVal Create As Boolean) As Object
    Set IFileSystem_OpenTextFile = fso.OpenTextFile(Path, Mode, Create)
End Function

' Implementa la verificaci?n de existencia de archivos
Private Function IFileSystem_FileExists(ByVal Path As String) As Boolean
    IFileSystem_FileExists = fso.FileExists(Path)
End Function

' Implementa la verificaci?n de existencia de carpetas
Private Function IFileSystem_FolderExists(ByVal Path As String) As Boolean
    IFileSystem_FolderExists = fso.FolderExists(Path)
End Function

' Implementa la copia de archivos
Private Sub IFileSystem_CopyFile(ByVal Source As String, ByVal Destination As String)
    fso.CopyFile Source, Destination
End Sub

' Implementa la eliminaci?n de archivos
Private Sub IFileSystem_DeleteFile(ByVal Path As String, Optional ByVal Force As Boolean = False)
    fso.DeleteFile Path, Force
End Sub

' Implementa la creaci?n de carpetas
Private Sub IFileSystem_CreateFolder(ByVal Path As String)
    On Error GoTo ErrorHandler
    If Len(Path) = 0 Then Exit Sub
    Dim parent As String: parent = fso.GetParentFolderName(Path)
    If Len(parent) > 0 And Not fso.FolderExists(parent) Then
        IFileSystem_CreateFolder parent
    End If
    If Not fso.FolderExists(Path) Then fso.CreateFolder Path
    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, "CFileSystem.CreateFolder", _
              "No se pudo crear la carpeta: " & Path & " - " & Err.Description
End Sub

' Helper privada para creaci?n recursiva de carpetas
Private Sub CreateFolderRecursive(ByVal folderPath As String)
    On Error GoTo Fail
    If fso.FolderExists(folderPath) Then Exit Sub
    Dim parentPath As String: parentPath = fso.GetParentFolderName(folderPath)
    If Len(parentPath) > 0 And Not fso.FolderExists(parentPath) Then CreateFolderRecursive parentPath
    fso.CreateFolder folderPath
    Exit Sub
Fail:
    Err.Raise Err.Number, "CFileSystem.CreateFolderRecursive", Err.Description
End Sub

' Implementa la eliminaci?n de carpetas
Private Sub IFileSystem_DeleteFolder(ByVal Path As String)
    fso.DeleteFolder Path
End Sub

' Implementa la eliminaci?n recursiva de carpetas
Private Sub IFileSystem_DeleteFolderRecursive(ByVal Path As String)

    Dim subFolder As Object ' Folder
    Dim file As Object      ' File
    Dim folder As Object    ' Folder
    
    On Error GoTo ErrorHandler ' Manejo de errores centralizado
    
    If Not fso.FolderExists(Path) Then Exit Sub
    
    Set folder = fso.GetFolder(Path)
    
    ' 1. Borrar todos los ficheros dentro de la carpeta actual
    For Each file In folder.Files
        file.Delete True ' True = forzar borrado si es de solo lectura
    Next file
    
    ' 2. Llamar recursivamente para cada subcarpeta
    For Each subFolder In folder.SubFolders
        Call IFileSystem_DeleteFolderRecursive(subFolder.Path)
    Next subFolder
    
    ' 3. Borrar la carpeta actual, que ahora est? vac?a
    folder.Delete True
    
    Exit Sub
    
ErrorHandler:
    ' No enmascarar el error, simplemente salir. El On Error del llamador lo gestionar?.
    ' Esto nos permitir? ver si hay un error de "Permiso denegado" por un fichero bloqueado.
End Sub

' Implementa la escritura de l?neas a archivo
Private Sub IFileSystem_WriteLineToFile(ByVal filePath As String, ByVal line As String)
    ' Encapsula la l?gica de abrir, escribir y cerrar.
    Dim file As Object
    ' 8 = ForAppending, True = Create if not exists
    Set file = fso.OpenTextFile(filePath, 8, True)
    file.WriteLine line
    file.Close
    Set file = Nothing
End Sub

' ============================================================================
' M?TODOS P?BLICOS DE CONVENIENCIA (Lecci?n 24)
' ============================================================================

Public Function OpenTextFile(ByVal Path As String, ByVal Mode As Long, ByVal Create As Boolean) As Object
    Set OpenTextFile = IFileSystem_OpenTextFile(Path, Mode, Create)
End Function

Public Function FileExists(ByVal Path As String) As Boolean
    FileExists = IFileSystem_FileExists(Path)
End Function

Public Function FolderExists(ByVal Path As String) As Boolean
    FolderExists = IFileSystem_FolderExists(Path)
End Function

Public Sub CopyFile(ByVal Source As String, ByVal Destination As String)
    IFileSystem_CopyFile Source, Destination
End Sub

Public Sub DeleteFile(ByVal Path As String, Optional ByVal Force As Boolean = False)
    IFileSystem_DeleteFile Path, Force
End Sub

Public Sub CreateFolder(ByVal Path As String)
    IFileSystem_CreateFolder Path
End Sub

Public Sub DeleteFolder(ByVal Path As String)
    IFileSystem_DeleteFolder Path
End Sub

Public Sub DeleteFolderRecursive(ByVal Path As String)
    ' Delega la llamada a la implementaci?n privada y segura de la interfaz
    Call IFileSystem_DeleteFolderRecursive(Path)
End Sub

Public Sub WriteLineToFile(ByVal filePath As String, ByVal line As String)
    Call IFileSystem_WriteLineToFile(filePath, line)
End Sub

