VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockDocumentService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IDocumentService

Private m_GenerarDocumento_ReturnValue As String
Private m_LeerDocumento_ReturnValue As ESolicitud
Private m_GenerarDocumento_WasCalled As Boolean
Private m_LeerDocumento_WasCalled As Boolean
Private m_GenerarDocumento_LastSolicitudId As Long
Private m_LeerDocumento_LastRuta As String

Private Sub Class_Initialize()
    Call Reset
End Sub

Public Sub Reset()
    m_GenerarDocumento_ReturnValue = ""
    Set m_LeerDocumento_ReturnValue = Nothing
    m_GenerarDocumento_WasCalled = False
    m_LeerDocumento_WasCalled = False
    m_GenerarDocumento_LastSolicitudId = 0
    m_LeerDocumento_LastRuta = ""
End Sub

Public Sub ConfigureGenerarDocumento(ByVal returnValue As String)
    m_GenerarDocumento_ReturnValue = returnValue
End Sub

Public Sub ConfigureLeerDocumento(ByVal returnValue As ESolicitud)
    Set m_LeerDocumento_ReturnValue = returnValue
End Sub

Public Property Get GenerarDocumento_WasCalled() As Boolean
    GenerarDocumento_WasCalled = m_GenerarDocumento_WasCalled
End Property

Public Property Get LeerDocumento_WasCalled() As Boolean
    LeerDocumento_WasCalled = m_LeerDocumento_WasCalled
End Property

Public Property Get GenerarDocumento_LastSolicitudId() As Long
    GenerarDocumento_LastSolicitudId = m_GenerarDocumento_LastSolicitudId
End Property

Public Property Get LeerDocumento_LastRuta() As String
    LeerDocumento_LastRuta = m_LeerDocumento_LastRuta
End Property

Private Function IDocumentService_GenerarDocumento(ByVal SolicitudID As Long) As String
    m_GenerarDocumento_WasCalled = True
    m_GenerarDocumento_LastSolicitudId = SolicitudID
    IDocumentService_GenerarDocumento = m_GenerarDocumento_ReturnValue
End Function

Private Function IDocumentService_LeerDocumento(ByVal rutaDocumento As String) As ESolicitud
    m_LeerDocumento_WasCalled = True
    m_LeerDocumento_LastRuta = rutaDocumento
    Set IDocumentService_LeerDocumento = m_LeerDocumento_ReturnValue
End Function
