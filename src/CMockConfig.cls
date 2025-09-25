VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit
'hola2222¡
Implements IConfig

Private m_Settings As Object

Private Sub Class_Initialize(): Set m_Settings = New Scripting.Dictionary: m_Settings.CompareMode = TextCompare: End Sub
Public Sub Reset(): Set m_Settings = New Scripting.Dictionary: m_Settings.CompareMode = TextCompare: End Sub

Private Function IConfig_GetValue(ByVal key As String) As Variant
    If m_Settings.exists(key) Then IConfig_GetValue = m_Settings(key) Else IConfig_GetValue = ""
End Function
Private Sub IConfig_SetSetting(ByVal key As String, ByVal value As Variant)
    If m_Settings.exists(key) Then m_Settings(key) = value Else m_Settings.Add key, value
End Sub

' --- Implementaciones Espec?ficas ---
Private Function IConfig_GetCondorDataPath() As String: IConfig_GetCondorDataPath = IConfig_GetValue("CONDOR_DATA_PATH"): End Function
Private Function IConfig_GetLanzaderaDataPath() As String: IConfig_GetLanzaderaDataPath = IConfig_GetValue("LANZADERA_DATA_PATH"): End Function
Private Function IConfig_GetExpedientesDataPath() As String: IConfig_GetExpedientesDataPath = IConfig_GetValue("EXPEDIENTES_DATA_PATH"): End Function
Private Function IConfig_GetCorreosDBPath() As String: IConfig_GetCorreosDBPath = IConfig_GetValue("CORREOS_DB_PATH"): End Function

Private Function IConfig_GetAttachmentsPath() As String: IConfig_GetAttachmentsPath = IConfig_GetValue("ATTACHMENTS_PATH"): End Function

' --- Contrase?as Espec?ficas ---
Private Function IConfig_GetCondorPassword() As String: IConfig_GetCondorPassword = IConfig_GetValue("CONDOR_PASSWORD"): End Function
Private Function IConfig_GetLanzaderaPassword() As String: IConfig_GetLanzaderaPassword = IConfig_GetValue("LANZADERA_PASSWORD"): End Function
Private Function IConfig_GetExpedientesPassword() As String: IConfig_GetExpedientesPassword = IConfig_GetValue("EXPEDIENTES_PASSWORD"): End Function
Private Function IConfig_GetCorreosPassword() As String: IConfig_GetCorreosPassword = IConfig_GetValue("CORREOS_PASSWORD"): End Function

' --- Configuraci?n General ---
Private Function IConfig_GetUsuarioActual() As String: IConfig_GetUsuarioActual = IConfig_GetValue("USUARIO_ACTUAL"): End Function
Private Function IConfig_GetCorreoAdministrador() As String: IConfig_GetCorreoAdministrador = IConfig_GetValue("CORREO_ADMINISTRADOR"): End Function
Private Function IConfig_GetIDAplicacionCondor() As String: IConfig_GetIDAplicacionCondor = IConfig_GetValue("ID_APLICACION_CONDOR"): End Function

' --- M?todos P?blicos de Conveniencia ---
Public Sub SetSetting(ByVal key As String, ByVal value As Variant)
    Call IConfig_SetSetting(key, value)
End Sub

Public Function GetValue(ByVal key As String) As Variant
    GetValue = IConfig_GetValue(key)
End Function

Public Function HasKey(ByVal key As String) As Boolean
    HasKey = m_Settings.exists(key)
End Function

Public Function GetAllKeys() As Variant
    GetAllKeys = m_Settings.Keys
End Function

