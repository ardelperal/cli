VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CDocumentService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IDocumentService

' Dependencias inyectadas
Private m_Config As IConfig
Private m_fileSystem As IFileSystem
Private m_WordManager As IWordManager
Private m_ErrorHandler As IErrorHandlerService
Private m_SolicitudService As ISolicitudService
Private m_MapeoRepo As IMapeoRepository

Public Sub Initialize(ByVal config As IConfig, ByVal fileSystem As IFileSystem, ByVal wordManager As IWordManager, ByVal ErrorHandler As IErrorHandlerService, ByVal solicitudService As ISolicitudService, ByVal mapeoRepo As IMapeoRepository)
    Set m_Config = config
    Set m_fileSystem = fileSystem
    Set m_WordManager = wordManager
    Set m_ErrorHandler = ErrorHandler
    Set m_SolicitudService = solicitudService
    Set m_MapeoRepo = mapeoRepo
End Sub

Private Function IDocumentService_GenerarDocumento(ByVal SolicitudID As Long) As String
    On Error GoTo ErrorHandler
    
    Dim solicitud As ESolicitud
    Dim mapeos As Object ' Scripting.Dictionary
    Dim key As Variant
    Dim plantillaPath As String, generatedDocsPath As String, generatedFilename As String, finalDocPath As String

    ' 1. Obtener datos de la solicitud
    Set solicitud = m_SolicitudService.ObtenerSolicitudPorId(SolicitudID)
    If solicitud Is Nothing Then
        m_ErrorHandler.LogError 404, "Solicitud no encontrada con ID: " & SolicitudID, "CDocumentService"
        Exit Function
    End If

    ' 2. Obtener mapeos para el tipo de solicitud
    Set mapeos = m_MapeoRepo.ObtenerMapeosPorCategoria(solicitud.tipoSolicitud)

    ' 3. Construir rutas din?micamente desde la configuraci?n
    Dim templateFilenameKey As String
    templateFilenameKey = "TEMPLATE_" & solicitud.tipoSolicitud & "_FILENAME"
    
    plantillaPath = JoinPath(m_Config.GetValue("TEMPLATES_PATH"), m_Config.GetValue(templateFilenameKey))
    generatedDocsPath = m_Config.GetValue("GENERATED_DOCS_PATH")
    generatedFilename = "Generated_" & solicitud.codigoSolicitud & "_" & Format(Now, "yyyymmddhhmmss") & ".docx"
    finalDocPath = JoinPath(generatedDocsPath, generatedFilename)

    ' 4. Asegurarse de que la carpeta de destino existe
    m_fileSystem.CreateFolder generatedDocsPath

    ' 5. Aprovisionar: Copiar plantilla a destino
     m_fileSystem.CopyFile plantillaPath, finalDocPath
    ' 6. Rellenar el documento
    If Not m_WordManager.AbrirDocumento(finalDocPath) Then
        m_ErrorHandler.LogError 500, "No se pudo abrir el documento generado.", "CDocumentService"
        Exit Function
    End If

    ' L?gica de relleno (simplificada para el ejemplo)
    ' En un caso real, aqu? ir?a el bucle sobre los mapeos
    m_WordManager.SetBookmarkText "BookmarkCodigo", solicitud.codigoSolicitud
    m_WordManager.SetBookmarkText "BookmarkUsuario", solicitud.usuarioCreacion

    m_WordManager.GuardarDocumento
    m_WordManager.Dispose

    ' 7. Devolver ruta real
    IDocumentService_GenerarDocumento = finalDocPath
    Exit Function

ErrorHandler:
    m_ErrorHandler.LogError Err.Number, Err.Description, "CDocumentService.GenerarDocumento"
    IDocumentService_GenerarDocumento = ""
End Function

Private Function IDocumentService_LeerDocumento(ByVal rutaDocumento As String) As ESolicitud
    ' Implementaci?n placeholder
    Set IDocumentService_LeerDocumento = New ESolicitud
End Function

Public Function GenerarDocumento(ByVal SolicitudID As Long) As String
    GenerarDocumento = IDocumentService_GenerarDocumento(SolicitudID)
End Function

Public Function LeerDocumento(ByVal rutaDocumento As String) As ESolicitud
    Set LeerDocumento = IDocumentService_LeerDocumento(rutaDocumento)
End Function


Private Function JoinPath(ByVal basePath As String, ByVal relativePath As String) As String
    Dim b As String: b = Trim$(basePath)
    Dim r As String: r = Trim$(relativePath)
    If Len(b) = 0 Then JoinPath = r: Exit Function
    If Len(r) = 0 Then JoinPath = b: Exit Function
    If Right$(b, 1) = "\" Then b = Left$(b, Len(b) - 1)
    If Left$(r, 1) = "\" Then r = Mid$(r, 2)
    JoinPath = b & "\" & r
End Function
