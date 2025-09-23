VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IFileSystem

Public WriteLineToFile_WasCalled As Boolean
Public WriteLineToFile_LastPath As String
Public WriteLineToFile_LastLine As String



Public Sub Reset()
    WriteLineToFile_WasCalled = False
    WriteLineToFile_LastPath = ""
    WriteLineToFile_LastLine = ""
End Sub



Private Sub IFileSystem_CreateFolder(ByVal Path As String): End Sub

Private Sub IFileSystem_CopyFile(ByVal Source As String, ByVal Destination As String): End Sub

Private Sub IFileSystem_DeleteFile(ByVal Path As String, Optional ByVal Force As Boolean = False): End Sub

Private Sub IFileSystem_DeleteFolder(ByVal Path As String): End Sub

Private Sub IFileSystem_DeleteFolderRecursive(ByVal Path As String): End Sub

Private Function IFileSystem_FileExists(ByVal Path As String) As Boolean: End Function

Private Function IFileSystem_FolderExists(ByVal Path As String) As Boolean: End Function

Private Function IFileSystem_OpenTextFile(ByVal Path As String, ByVal Mode As Long, ByVal Create As Boolean) As Object: End Function

Private Sub IFileSystem_WriteLineToFile(ByVal filePath As String, ByVal line As String)
    WriteLineToFile_WasCalled = True
    WriteLineToFile_LastPath = filePath
    WriteLineToFile_LastLine = line
End Sub
