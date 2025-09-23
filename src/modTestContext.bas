Attribute VB_Name = "modTestContext"
Option Compare Database
Option Explicit


' =================================================================
' MÓDULO: modTestContext
' DESCRIPCIÓN: Gestor Singleton para la configuración de pruebas.
' Versión: 2.1 - Restaurado aprovisionamiento de BBDD (Lección 55)
' =================================================================

Private g_TestConfig As IConfig

Public Function GetTestConfig() As IConfig
    If g_TestConfig Is Nothing Then
        Dim realConfig As New CConfig
        realConfig.LoadConfiguration
        
        Dim mockConfigImpl As New CMockConfig
        
        ' --- Definición de Rutas de ORIGEN (Masters en /back) ---
        mockConfigImpl.SetSetting "DEV_CONDOR_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetProjectPath(), "back\data\CONDOR_datos.accdb")
        mockConfigImpl.SetSetting "DEV_LANZADERA_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetProjectPath(), "back\data\Lanzadera_Datos.accdb")
        mockConfigImpl.SetSetting "DEV_EXPEDIENTES_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetProjectPath(), "back\data\Expedientes_datos.accdb")
        mockConfigImpl.SetSetting "DEV_CORREOS_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetProjectPath(), "back\data\correos_datos.accdb")

        ' --- Definición de Rutas de DESTINO (Copias temporales en /workspace) ---
        mockConfigImpl.SetSetting "CONDOR_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "CONDOR_integration_test.accdb")
        mockConfigImpl.SetSetting "LANZADERA_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "Lanzadera_integration_test.accdb")
        mockConfigImpl.SetSetting "EXPEDIENTES_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "Expedientes_integration_test.accdb")
        mockConfigImpl.SetSetting "CORREOS_DATA_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "Correos_integration_test.accdb")
        mockConfigImpl.SetSetting "CORREOS_DB_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "Correos_integration_test.accdb") ' Clave duplicada por compatibilidad
        
        ' --- AÑADIDAS: Rutas de workspace para plantillas y documentos generados ---
        mockConfigImpl.SetSetting "TEMPLATES_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "test_templates\")
        mockConfigImpl.SetSetting "GENERATED_DOCS_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "generated_documents\")
        
        ' --- Descubrimiento dinámico de configuración real ---
        mockConfigImpl.SetSetting "PRODUCTION_TEMPLATES_PATH", realConfig.GetValue("TEMPLATES_PATH")
        mockConfigImpl.SetSetting "TEMPLATE_PC_FILENAME", realConfig.GetValue("TEMPLATE_NAME_PC")
        mockConfigImpl.SetSetting "TEMPLATE_CDCA_FILENAME", realConfig.GetValue("TEMPLATE_NAME_CDCA")
        mockConfigImpl.SetSetting "TEMPLATE_CDCASUB_FILENAME", realConfig.GetValue("TEMPLATE_NAME_CDCASUB")
        mockConfigImpl.SetSetting "ID_APLICACION_CONDOR", realConfig.GetValue("ID_APLICACION_CONDOR")
        
        ' --- Parámetros específicos de prueba ---
        mockConfigImpl.SetSetting "LOG_FILE_PATH", modTestUtils.JoinPath(modTestUtils.GetWorkspacePath(), "tests.log")
        mockConfigImpl.SetSetting "USUARIO_ACTUAL", "test.user@condor.com"
        mockConfigImpl.SetSetting "LANZADERA_PASSWORD", "dpddpd"
        mockConfigImpl.SetSetting "EXPEDIENTES_PASSWORD", "dpddpd"
        mockConfigImpl.SetSetting "CORREOS_PASSWORD", "dpddpd"
        mockConfigImpl.SetSetting "CONDOR_PASSWORD", ""
        
        Set g_TestConfig = mockConfigImpl
    End If
    Set GetTestConfig = g_TestConfig
End Function
Public Sub ResetTestContext()
    Set g_TestConfig = Nothing
End Sub
