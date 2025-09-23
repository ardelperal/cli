VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

    Implements IConfig

    Private m_Settings As Object

    Private Sub Class_Initialize()
        Set m_Settings = New Scripting.Dictionary
        m_Settings.CompareMode = TextCompare ' Case-insensitive keys
    End Sub

    Public Sub LoadConfiguration()
        On Error GoTo ErrorHandler
        
        Dim db As DAO.Database, rs As DAO.Recordset
        
        ' 1. LEER EL CONMUTADOR DE ENTORNO DESDE EL FRONTEND
        Set db = CurrentDb
        Set rs = db.OpenRecordset("SELECT TOP 1 Entorno FROM TbLocalConfig")
        
        Dim entorno As String
        If Not rs.EOF Then
            entorno = UCase(Trim(rs.Fields("Entorno").value & ""))
        Else
            entorno = "DESARROLLO" ' Fallback seguro
        End If
        rs.Close
        
        ' 2. ESTABLECER RUTAS POR DEFECTO BASADAS EN EL ENTORNO
        If entorno = "PRODUCCION" Then
            Dim prodBasePath As String: prodBasePath = "\\datoste\APLICACIONES_DYS\Aplicaciones PpD\"
            m_Settings("DATA_PATH") = JoinPath(prodBasePath, "Condor\condor_datos.accdb")
            m_Settings("LANZADERA_DATA_PATH") = JoinPath(prodBasePath, "0Lanzadera\Lanzadera_Datos.accdb")
            m_Settings("EXPEDIENTES_DATA_PATH") = JoinPath(prodBasePath, "EXPEDIENTES\Expedientes_datos.accdb")
            m_Settings("CORREOS_DATA_PATH") = JoinPath(prodBasePath, "00Recursos\correos_datos.accdb") ' CORREGIDO
            m_Settings("TEMPLATES_PATH") = JoinPath(JoinPath(prodBasePath, "Condor\"), "recursos\Plantillas\")
        Else ' DESARROLLO
            Dim devBasePath As String: devBasePath = modTestUtils.GetProjectPath()
            Dim backData As String: backData = JoinPath(devBasePath, "back\data")
            Dim frontRoot As String: frontRoot = JoinPath(devBasePath, "front")
            
            m_Settings("DATA_PATH") = JoinPath(backData, "CONDOR_datos.accdb")
            m_Settings("LANZADERA_DATA_PATH") = JoinPath(backData, "Lanzadera_Datos.accdb")
            m_Settings("EXPEDIENTES_DATA_PATH") = JoinPath(backData, "Expedientes_datos.accdb")
            m_Settings("CORREOS_DATA_PATH") = JoinPath(backData, "correos_datos.accdb")
            m_Settings("TEMPLATES_PATH") = JoinPath(frontRoot, "recursos\Plantillas\")
        End If

        ' 3. LEER *TODA* LA CONFIGURACI?N DEL BACKEND PARA SOBRESCRIBIR
        Dim backendPath As String: backendPath = m_Settings("DATA_PATH")
        
        Dim backendPassword As String: backendPassword = "dpddpd"
        If entorno = "DESARROLLO" Then backendPassword = ""

        Set db = DBEngine.OpenDatabase(backendPath, False, True, ";PWD=" & backendPassword)
        Set rs = db.OpenRecordset("SELECT clave, valor FROM tbConfiguracion")
        
        Do While Not rs.EOF
            If Not IsNull(rs!clave.value) And Len(Trim(rs!clave.value & "")) > 0 Then
                m_Settings(rs!clave.value) = rs!valor.value
            End If
            rs.MoveNext
        Loop
        rs.Close
        db.Close
        
        Exit Sub
ErrorHandler:
        If Not rs Is Nothing Then rs.Close
        If Not db Is Nothing Then db.Close
        Err.Raise Err.Number, "CConfig.LoadConfiguration", Err.Description
    End Sub

    Public Sub LoadFromDictionary(ByVal settings As Object)
        Set m_Settings = settings
    End Sub

    Private Function IConfig_GetValue(ByVal key As String) As Variant
        If m_Settings.Exists(key) Then ' CORREGIDO
            IConfig_GetValue = m_Settings(key)
        Else
            IConfig_GetValue = ""
        End If
    End Function

    Private Sub IConfig_SetSetting(ByVal key As String, ByVal value As Variant)
        If m_Settings.Exists(key) Then ' CORREGIDO
            m_Settings(key) = value
        Else
            m_Settings.Add key, value
        End If
    End Sub
        
    Private Function JoinPath(ByVal basePath As String, ByVal relativePath As String) As String
        Dim b As String: b = Trim$(basePath)
        Dim r As String: r = Trim$(relativePath)
        If Len(b) = 0 Then JoinPath = r: Exit Function
        If Len(r) = 0 Then JoinPath = b: Exit Function
        If Right$(b, 1) = "\" Then b = Left$(b, Len(b) - 1)
        If Left$(r, 1) = "\" Then r = Mid$(r, 2)
        JoinPath = b & "\" & r
    End Function

    Private Function IConfig_GetCorreosDBPath() As String
        IConfig_GetCorreosDBPath = IConfig_GetValue("CORREOS_DATA_PATH")
    End Function

    Private Function IConfig_GetUsuarioActual() As String
        IConfig_GetUsuarioActual = IConfig_GetValue("USUARIO_ACTUAL")
    End Function

    Private Function IConfig_GetCorreoAdministrador() As String
        IConfig_GetCorreoAdministrador = IConfig_GetValue("CORREO_ADMINISTRADOR")
    End Function

    Private Function IConfig_GetIDAplicacionCondor() As String
        IConfig_GetIDAplicacionCondor = IConfig_GetValue("ID_APLICACION_CONDOR")
    End Function

    Private Function IConfig_GetLanzaderaDataPath() As String
        IConfig_GetLanzaderaDataPath = IConfig_GetValue("LANZADERA_DATA_PATH")
    End Function

    Private Function IConfig_GetLanzaderaPassword() As String
        IConfig_GetLanzaderaPassword = IConfig_GetValue("LANZADERA_PASSWORD")
    End Function

    Private Function IConfig_GetCondorDataPath() As String
        IConfig_GetCondorDataPath = IConfig_GetValue("CONDOR_DATA_PATH")
    End Function

    Private Function IConfig_GetExpedientesDataPath() As String
        IConfig_GetExpedientesDataPath = IConfig_GetValue("EXPEDIENTES_DATA_PATH")
    End Function

    Private Function IConfig_GetCondorPassword() As String
        IConfig_GetCondorPassword = IConfig_GetValue("CONDOR_PASSWORD")
    End Function

    Private Function IConfig_GetExpedientesPassword() As String
        IConfig_GetExpedientesPassword = IConfig_GetValue("EXPEDIENTES_PASSWORD")
    End Function

    Private Function IConfig_GetCorreosPassword() As String
        IConfig_GetCorreosPassword = IConfig_GetValue("CORREOS_PASSWORD")
    End Function

    '==============================================================================
    ' M?TODOS P?BLICOS DE CONVENIENCIA (Lecci?n 24)
    '==============================================================================

    Public Function GetValue(ByVal key As String) As Variant
        GetValue = IConfig_GetValue(key)
    End Function

    Public Function GetCorreosDBPath() As String
        GetCorreosDBPath = IConfig_GetCorreosDBPath()
    End Function

    Public Function GetUsuarioActual() As String
        GetUsuarioActual = IConfig_GetUsuarioActual()
    End Function

    Public Function GetCorreoAdministrador() As String
        GetCorreoAdministrador = IConfig_GetCorreoAdministrador()
    End Function

    Public Function GetIDAplicacionCondor() As String
        GetIDAplicacionCondor = IConfig_GetIDAplicacionCondor()
    End Function

    Public Function GetLanzaderaDataPath() As String
        GetLanzaderaDataPath = IConfig_GetLanzaderaDataPath()
    End Function

    Public Function GetLanzaderaPassword() As String
        GetLanzaderaPassword = IConfig_GetLanzaderaPassword()
    End Function

    Public Function GetCondorDataPath() As String
        GetCondorDataPath = IConfig_GetCondorDataPath()
    End Function

    Public Function GetExpedientesDataPath() As String
        GetExpedientesDataPath = IConfig_GetExpedientesDataPath()
    End Function

    Public Function GetCondorPassword() As String
        GetCondorPassword = IConfig_GetCondorPassword()
    End Function

    Public Function GetExpedientesPassword() As String
        GetExpedientesPassword = IConfig_GetExpedientesPassword()
    End Function

    Public Function GetCorreosPassword() As String
        GetCorreosPassword = IConfig_GetCorreosPassword()
    End Function

    Public Sub SetSetting(ByVal key As String, ByVal value As Variant)
        Call IConfig_SetSetting(key, value)
    End Sub

    Public Function GetAllKeys() As Variant
        GetAllKeys = m_Settings.Keys
    End Function
