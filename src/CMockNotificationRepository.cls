VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockNotificationRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements INotificationRepository

' --- Variables "Esp?a" privadas ---
Private m_WasEnqueueEmailCalled As Boolean
Private m_LastDestinatarios As String
Private m_LastAsunto As String
Private m_LastCuerpoHTML As String
Private m_LastConCopia As String
Private m_LastConCopiaOculta As String
Private m_LastAdjuntoPath As String

' --- Variable de configuraci?n de comportamiento ---
Private m_EnqueueEmailReturnValue As Boolean

' Constructor
Private Sub Class_Initialize()
    Reset
End Sub

' Reinicia el estado del mock
Public Sub Reset()
    m_WasEnqueueEmailCalled = False
    m_LastDestinatarios = ""
    m_LastAsunto = ""
    m_LastCuerpoHTML = ""
    m_LastConCopia = ""
    m_LastConCopiaOculta = ""
    m_LastAdjuntoPath = ""
    m_EnqueueEmailReturnValue = True ' Comportamiento por defecto: ?xito
End Sub

' --- M?todo de configuraci?n ---
Public Sub ConfigureEnqueueEmail(ByVal returnValue As Boolean)
    m_EnqueueEmailReturnValue = returnValue
End Sub

' --- Propiedades "Esp?a" p?blicas de solo lectura ---
Public Property Get WasEnqueueEmailCalled() As Boolean
    WasEnqueueEmailCalled = m_WasEnqueueEmailCalled
End Property
Public Property Get LastDestinatarios() As String
    LastDestinatarios = m_LastDestinatarios
End Property
Public Property Get LastAsunto() As String
    LastAsunto = m_LastAsunto
End Property
Public Property Get LastCuerpoHTML() As String
    LastCuerpoHTML = m_LastCuerpoHTML
End Property
Public Property Get LastConCopia() As String
    LastConCopia = m_LastConCopia
End Property
Public Property Get LastConCopiaOculta() As String
    LastConCopiaOculta = m_LastConCopiaOculta
End Property
Public Property Get LastAdjuntoPath() As String
    LastAdjuntoPath = m_LastAdjuntoPath
End Property

' --- Implementaci?n de la interfaz ---
Private Function INotificationRepository_EnqueueEmail(ByVal destinatarios As String, ByVal asunto As String, ByVal cuerpoHTML As String, Optional ByVal conCopia As String = "", Optional ByVal conCopiaOculta As String = "", Optional ByVal adjuntoPath As String = "") As Boolean
    ' Registrar la llamada y los par?metros para verificaci?n
    m_WasEnqueueEmailCalled = True
    m_LastDestinatarios = destinatarios
    m_LastAsunto = asunto
    m_LastCuerpoHTML = cuerpoHTML
    m_LastConCopia = conCopia
    m_LastConCopiaOculta = conCopiaOculta
    m_LastAdjuntoPath = adjuntoPath
    
    ' Devolver el valor configurado
    INotificationRepository_EnqueueEmail = m_EnqueueEmailReturnValue
End Function
