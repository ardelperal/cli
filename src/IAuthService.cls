VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAuthService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Function GetUserRole(ByVal userEmail As String) As UserRole
End Function

Public Function AuthenticateUser(ByVal userEmail As String, ByVal password As String) As Boolean
End Function

Public Function GetCurrentUserEmail() As String
End Function
