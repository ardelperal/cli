VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockWordManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IWordManager

' --- Variables de configuraci?n de comportamiento ---
Private m_AbrirDocumento_ReturnValue As Boolean
Private m_GuardarDocumento_ReturnValue As Boolean
Private m_SetBookmarkText_ReturnValue As Boolean
Private m_GetBookmarkText_ReturnValue As String

' --- Variables "Esp?a" para verificaci?n ---
Private m_AbrirDocumento_WasCalled As Boolean
Private m_GuardarDocumento_WasCalled As Boolean
Private m_Dispose_WasCalled As Boolean

Private m_SetBookmarkText_WasCalled As Boolean
Private m_SetBookmarkText_LastBookmarkName As String
Private m_SetBookmarkText_LastValue As String

' Constructor
Private Sub Class_Initialize()
    Reset
End Sub

' --- M?todos p?blicos para configurar el mock ---
Public Sub ConfigureAbrirDocumento(ByVal returnValue As Boolean)
    m_AbrirDocumento_ReturnValue = returnValue
End Sub

Public Sub ConfigureGuardarDocumento(ByVal returnValue As Boolean)
    m_GuardarDocumento_ReturnValue = returnValue
End Sub

Public Sub ConfigureSetBookmarkText(ByVal returnValue As Boolean)
    m_SetBookmarkText_ReturnValue = returnValue
End Sub

Public Sub ConfigureGetBookmarkText(ByVal returnValue As String)
    m_GetBookmarkText_ReturnValue = returnValue
End Sub

' --- M?todos p?blicos "Esp?a" de solo lectura ---
Public Property Get AbrirDocumento_WasCalled() As Boolean
    AbrirDocumento_WasCalled = m_AbrirDocumento_WasCalled
End Property

Public Property Get GuardarDocumento_WasCalled() As Boolean
    GuardarDocumento_WasCalled = m_GuardarDocumento_WasCalled
End Property

Public Property Get Dispose_WasCalled() As Boolean
    Dispose_WasCalled = m_Dispose_WasCalled
End Property

Public Property Get SetBookmarkText_WasCalled() As Boolean
    SetBookmarkText_WasCalled = m_SetBookmarkText_WasCalled
End Property

Public Property Get SetBookmarkText_LastBookmarkName() As String
    SetBookmarkText_LastBookmarkName = m_SetBookmarkText_LastBookmarkName
End Property

Public Property Get SetBookmarkText_LastValue() As String
    SetBookmarkText_LastValue = m_SetBookmarkText_LastValue
End Property


Public Sub Reset()
    m_AbrirDocumento_ReturnValue = False
    m_GuardarDocumento_ReturnValue = False
    m_SetBookmarkText_ReturnValue = False
    m_GetBookmarkText_ReturnValue = ""
    
    m_AbrirDocumento_WasCalled = False
    m_GuardarDocumento_WasCalled = False
    m_Dispose_WasCalled = False
    
    m_SetBookmarkText_WasCalled = False
    m_SetBookmarkText_LastBookmarkName = ""
    m_SetBookmarkText_LastValue = ""
End Sub

' --- Implementaci?n de la Interfaz IWordManager ---

Private Function IWordManager_AbrirDocumento(ByVal rutaDocumento As String) As Boolean
    m_AbrirDocumento_WasCalled = True
    IWordManager_AbrirDocumento = m_AbrirDocumento_ReturnValue
End Function

Private Function IWordManager_GuardarDocumento() As Boolean
    m_GuardarDocumento_WasCalled = True
    IWordManager_GuardarDocumento = m_GuardarDocumento_ReturnValue
End Function

Private Sub IWordManager_Dispose()
    m_Dispose_WasCalled = True
End Sub

Private Function IWordManager_SetBookmarkText(ByVal BookmarkName As String, ByVal value As String) As Boolean
    m_SetBookmarkText_WasCalled = True
    m_SetBookmarkText_LastBookmarkName = BookmarkName
    m_SetBookmarkText_LastValue = value
    IWordManager_SetBookmarkText = m_SetBookmarkText_ReturnValue
End Function

Private Function IWordManager_GetBookmarkText(ByVal BookmarkName As String) As String
    IWordManager_GetBookmarkText = m_GetBookmarkText_ReturnValue
End Function
