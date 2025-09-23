VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IAppManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Sub Initialize(ByVal authService As IAuthService, ByVal config As IConfig, ByVal ErrorHandler As IErrorHandlerService)
End Sub

' =====================================================
' METODO: StartApplication
' PROPOSITO: Iniciar la aplicaci?n con autenticaci?n y configuraci?n
' PARAMETROS:
'   - userEmail: Email del usuario (opcional para testing)
' RETORNA: Boolean - True si la inicializaci?n fue exitosa
' =====================================================
Public Function StartApplication(Optional ByVal userEmail As String = "") As Boolean
End Function

' =====================================================
' METODO: GetCurrentUserRole
' PROPOSITO: Obtener el rol del usuario actual autenticado
' RETORNA: UserRole - Rol del usuario actual
' =====================================================
Public Function GetCurrentUserRole() As UserRole
End Function

' =====================================================
' METODO: IsApplicationInitialized
' PROPOSITO: Verificar si la aplicaci?n est? correctamente inicializada
' RETORNA: Boolean - True si est? inicializada
' =====================================================
Public Function IsApplicationInitialized() As Boolean
End Function

' =====================================================
' METODO: Ping
' PROPOSITO: Funci?n de diagn?stico para smoke test
' RETORNA: String - "Pong"
' =====================================================
Public Function Ping() As String
End Function
