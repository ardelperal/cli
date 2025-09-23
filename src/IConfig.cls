VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Function GetValue(ByVal key As String) As Variant
End Function

Public Sub SetSetting(ByVal key As String, ByVal value As Variant)
End Sub

' --- Rutas Espec?ficas ---
Public Function GetCondorDataPath() As String
End Function

Public Function GetLanzaderaDataPath() As String
End Function

Public Function GetExpedientesDataPath() As String
End Function

Public Function GetCorreosDBPath() As String
End Function

' --- Contrase?as Espec?ficas ---
Public Function GetCondorPassword() As String
End Function

Public Function GetLanzaderaPassword() As String
End Function

Public Function GetExpedientesPassword() As String
End Function

Public Function GetCorreosPassword() As String
End Function

' --- Configuraci?n General ---
Public Function GetUsuarioActual() As String
End Function

Public Function GetCorreoAdministrador() As String
End Function

Public Function GetIDAplicacionCondor() As String
End Function
