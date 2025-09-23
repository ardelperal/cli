VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFileSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Function OpenTextFile(ByVal Path As String, ByVal Mode As Long, ByVal Create As Boolean) As Object
End Function

' Verifica si un archivo existe
' @param Path: Ruta del archivo a verificar
' @return: True si el archivo existe, False en caso contrario
Public Function FileExists(ByVal Path As String) As Boolean
End Function

' Verifica si una carpeta existe
' @param Path: Ruta de la carpeta a verificar
' @return: True si la carpeta existe, False en caso contrario
Public Function FolderExists(ByVal Path As String) As Boolean
End Function

' Copia un archivo de origen a destino
' @param Source: Ruta del archivo origen
' @param Destination: Ruta del archivo destino
Public Sub CopyFile(ByVal Source As String, ByVal Destination As String)
End Sub

' Elimina un archivo
' @param Path: Ruta del archivo a eliminar
' @param Force: Si debe forzar la eliminaci?n (opcional, por defecto False)
Public Sub DeleteFile(ByVal Path As String, Optional ByVal Force As Boolean = False)
End Sub

' Crea una carpeta
' @param Path: Ruta de la carpeta a crear
Public Sub CreateFolder(ByVal Path As String)
End Sub

' Elimina una carpeta
' @param Path: Ruta de la carpeta a eliminar
Public Sub DeleteFolder(ByVal Path As String)
End Sub

' Elimina una carpeta de forma recursiva (incluyendo todo su contenido)
' @param Path: Ruta de la carpeta a eliminar recursivamente
Public Sub DeleteFolderRecursive(ByVal Path As String)
End Sub

' Escribe una l?nea completa en un fichero, manejando la apertura y el cierre internamente.
' @param filePath: Ruta completa al archivo.
' @param line: El texto que se va a escribir en el archivo.
Public Sub WriteLineToFile(ByVal filePath As String, ByVal line As String)
End Sub
