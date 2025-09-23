VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockAuthService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IAuthService

Private m_GetUserRole_Result As UserRole
Private m_AuthenticateUser_Result As Boolean
Private m_GetCurrentUserEmail_Result As String

Private Sub Class_Initialize(): Call Reset: End Sub
Public Sub Reset(): m_GetUserRole_Result = RolDesconocido: m_AuthenticateUser_Result = False: m_GetCurrentUserEmail_Result = "": End Sub

Public Sub ConfigureGetUserRole(ByVal value As UserRole): m_GetUserRole_Result = value: End Sub
Public Sub ConfigureAuthenticateUser(ByVal value As Boolean): m_AuthenticateUser_Result = value: End Sub
Public Sub ConfigureGetCurrentUserEmail(ByVal value As String): m_GetCurrentUserEmail_Result = value: End Sub

Private Function IAuthService_GetUserRole(ByVal userEmail As String) As UserRole
    IAuthService_GetUserRole = m_GetUserRole_Result
End Function

Private Function IAuthService_AuthenticateUser(ByVal userEmail As String, ByVal password As String) As Boolean
    IAuthService_AuthenticateUser = m_AuthenticateUser_Result
End Function

Private Function IAuthService_GetCurrentUserEmail() As String
    IAuthService_GetCurrentUserEmail = m_GetCurrentUserEmail_Result
End Function

Public Function GetCurrentUserEmail() As String
    GetCurrentUserEmail = IAuthService_GetCurrentUserEmail()
End Function
