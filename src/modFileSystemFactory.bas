Attribute VB_Name = "modFileSystemFactory"
Option Compare Database
Option Explicit


' Módulo: modFileSystemFactory
' Descripción: Factory para crear instancias de IFileSystem
' Arquitectura: Capa de Servicios - Factory Pattern

Public Function CreateFileSystem(Optional ByVal config As IConfig = Nothing) As IFileSystem
    On Error GoTo ErrorHandler
    
    ' 1. Obtener la configuración efectiva
    Dim effectiveConfig As IConfig
    If config Is Nothing Then
        ' Si no se pasa una configuración, usar la global por defecto
        Set effectiveConfig = modTestContext.GetTestConfig()
    Else
        ' Si se pasa una configuración (desde un test), usarla
        Set effectiveConfig = config
    End If
    
    ' 2. Crear la instancia de la clase
    Dim fileSystemInstance As New CFileSystem
    
    ' 3. INICIALIZAR la instancia con sus dependencias
    fileSystemInstance.Initialize effectiveConfig
    
    ' 4. Devolver el objeto listo para usar
    Set CreateFileSystem = fileSystemInstance
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error en modFileSystemFactory.CreateFileSystem: " & Err.Number & " - " & Err.Description
    Set CreateFileSystem = Nothing
End Function
