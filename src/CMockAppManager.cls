VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockAppManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IAppManager

Private m_StartApplication_Result As Boolean
Private m_GetCurrentUserRole_Result As UserRole
Private m_IsInitialized_Result As Boolean
Private m_Ping_Result As String

Private m_StartApplication_WasCalled As Boolean

Private Sub Class_Initialize(): Call Reset: End Sub

Public Sub Reset()
    m_StartApplication_Result = True
    m_GetCurrentUserRole_Result = RolDesconocido
    m_IsInitialized_Result = True
    m_Ping_Result = "Pong"
    m_StartApplication_WasCalled = False
End Sub

Public Property Get StartApplication_WasCalled() As Boolean
    StartApplication_WasCalled = m_StartApplication_WasCalled
End Property

Public Sub ConfigureStartApplication(ByVal value As Boolean)
    m_StartApplication_Result = value
End Sub

Public Sub ConfigureGetCurrentUserRole(ByVal value As UserRole)
    m_GetCurrentUserRole_Result = value
End Sub

Private Sub IAppManager_Initialize(ByVal authService As IAuthService, ByVal config As IConfig, ByVal ErrorHandler As IErrorHandlerService)
    ' En el mock, la inicializaci?n no necesita hacer nada.
End Sub

Private Function IAppManager_StartApplication(Optional ByVal userEmail As String = "") As Boolean
    m_StartApplication_WasCalled = True
    IAppManager_StartApplication = m_StartApplication_Result
End Function

Private Function IAppManager_GetCurrentUserRole() As UserRole
    IAppManager_GetCurrentUserRole = m_GetCurrentUserRole_Result
End Function

Private Function IAppManager_IsApplicationInitialized() As Boolean
    IAppManager_IsApplicationInitialized = m_IsInitialized_Result
End Function

Private Function IAppManager_Ping() As String
    IAppManager_Ping = m_Ping_Result
End Function
