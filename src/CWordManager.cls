VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWordManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IWordManager

' Variables privadas para manejo de Word
Private m_WordApp As Object
Private m_WordDoc As Object
Private m_IsWordAppCreated As Boolean
Private m_ErrorHandler As IErrorHandlerService ' Inyectar ErrorHandler

' ============================================================================
' CONSTRUCTOR Y DESTRUCTOR
' ============================================================================

Private Sub Class_Initialize()
    m_IsWordAppCreated = False
    Set m_WordApp = Nothing
    Set m_WordDoc = Nothing
End Sub

Private Sub Class_Terminate()
    Call LimpiarRecursos
End Sub

' ============================================================================
' INICIALIZACI?N CON DEPENDENCIAS
' ============================================================================
Public Sub Initialize(ByVal wordApp As Object, ByVal ErrorHandler As IErrorHandlerService)
    Set m_WordApp = wordApp
    Set m_ErrorHandler = ErrorHandler
    m_IsWordAppCreated = False ' La aplicaci?n Word se inyecta, no se crea internamente
End Sub

' ============================================================================
' IMPLEMENTACI?N DE LA INTERFAZ IWordManager
' ============================================================================

Private Function IWordManager_AbrirDocumento(ByVal rutaDocumento As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Verificar que la aplicaci?n Word est? disponible
    If m_WordApp Is Nothing Then
        If Not m_ErrorHandler Is Nothing Then
            m_ErrorHandler.LogError -1, "Aplicaci?n Word no inicializada", "CWordManager.IWordManager_AbrirDocumento"
        End If
        IWordManager_AbrirDocumento = False
        Exit Function
    End If
    
    ' Abrir el documento
    Set m_WordDoc = m_WordApp.Documents.Open(rutaDocumento)
    
    IWordManager_AbrirDocumento = True
    Exit Function
    
ErrorHandler:
    ' Si el fichero no se encuentra (error com?n), no lo logueamos como un error del sistema.
    ' La funci?n simplemente devolver? False, que es el comportamiento esperado.
    IWordManager_AbrirDocumento = False
End Function

Private Function IWordManager_SetBookmarkText(ByVal BookmarkName As String, ByVal value As String) As Boolean
    On Error GoTo ErrorHandler

    IWordManager_SetBookmarkText = False
    If m_WordDoc Is Nothing Then Exit Function

    If m_WordDoc.Bookmarks.Exists(BookmarkName) Then
        m_WordDoc.FormFields(BookmarkName).result = value
        IWordManager_SetBookmarkText = True
    End If
    
    Exit Function
ErrorHandler:
    ' No se considera un error del sistema si el bookmark no existe.
    ' Simplemente se devuelve False.
End Function

Private Function IWordManager_GetBookmarkText(ByVal BookmarkName As String) As String
    On Error GoTo ErrorHandler
    
    IWordManager_GetBookmarkText = ""
    If m_WordDoc Is Nothing Then Exit Function

    If m_WordDoc.Bookmarks.Exists(BookmarkName) Then
        IWordManager_GetBookmarkText = m_WordDoc.FormFields(BookmarkName).result
    End If

    Exit Function
ErrorHandler:
    IWordManager_GetBookmarkText = ""
End Function

Private Function IWordManager_GuardarDocumento() As Boolean
    On Error GoTo ErrorHandler
    If m_WordDoc Is Nothing Then Exit Function
    
    m_WordDoc.Save
    IWordManager_GuardarDocumento = True
    
    Exit Function
ErrorHandler:
    IWordManager_GuardarDocumento = False
    ' El error ya ser? logueado por el llamador si es necesario.
End Function

' ============================================================================
' M?TODOS P?BLICOS DE CONVENIENCIA (Lecci?n 24)
' ============================================================================

Public Function AbrirDocumento(ByVal rutaDocumento As String) As Boolean
    AbrirDocumento = IWordManager_AbrirDocumento(rutaDocumento)
End Function

Public Function SetBookmarkText(ByVal BookmarkName As String, ByVal value As String) As Boolean
    SetBookmarkText = IWordManager_SetBookmarkText(BookmarkName, value)
End Function

Public Function GetBookmarkText(ByVal BookmarkName As String) As String
    GetBookmarkText = IWordManager_GetBookmarkText(BookmarkName)
End Function

Public Function GuardarDocumento() As Boolean
    GuardarDocumento = IWordManager_GuardarDocumento()
End Function

' ============================================================================
' M?TODOS PRIVADOS
' ============================================================================

Private Sub LimpiarRecursos()
    On Error Resume Next
    
    If Not m_WordDoc Is Nothing Then
        m_WordDoc.Close False
        Set m_WordDoc = Nothing
    End If
    
    If m_IsWordAppCreated And Not m_WordApp Is Nothing Then
        m_WordApp.Quit
        Set m_WordApp = Nothing
        m_IsWordAppCreated = False
    End If
End Sub

Private Sub IWordManager_Dispose()
    On Error Resume Next
    
    ' 1. Cerrar documento actual si existe
    If Not m_WordDoc Is Nothing Then
        m_WordDoc.Close False
        Set m_WordDoc = Nothing
    End If
    
    ' 2. Cerrar todos los documentos abiertos y la aplicaci?n Word
    If Not m_WordApp Is Nothing Then
        Do While m_WordApp.Documents.Count > 0
            m_WordApp.Documents(1).Close False
        Loop
        m_WordApp.Quit 0
        Set m_WordApp = Nothing
    End If
    
    ' 3. Resetear flags de estado
    m_IsWordAppCreated = False
End Sub
