VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMapeoRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IMapeoRepository

Private m_Config As IConfig
Private m_ErrorHandler As IErrorHandlerService

Public Sub Initialize(ByVal configService As IConfig, ByVal ErrorHandler As IErrorHandlerService)
    Set m_Config = configService
    Set m_ErrorHandler = ErrorHandler
End Sub

Private Function IMapeoRepository_GetMapeoPorTipo(ByVal tipoSolicitud As String) As EMapeo
    On Error GoTo ErrorHandler
    
    ' Variables para la consulta parametrizada
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    Dim params As Collection
    Dim prm As EQueryParameter
    Dim mapeo As EMapeo
    
    Set db = DBEngine.OpenDatabase(m_Config.GetCondorDataPath(), dbFailOnError, False, "MS Access;PWD=" & m_Config.GetCondorPassword())
    
    ' Crear QueryDef con consulta desde modQueries
    Set qdf = db.CreateQueryDef("", "SELECT * FROM tbMapeoCampos WHERE NombrePlantilla = [pNombrePlantilla]")
    
    ' Crear colecci?n de par?metros usando el tipoSolicitud
    Set params = New Collection
    Set prm = New EQueryParameter
    prm.ParameterName = "pNombrePlantilla"
    prm.ParameterValue = tipoSolicitud
    params.Add prm
    
    ' Asignar par?metros desde la colecci?n
    For Each prm In params
        qdf.Parameters(prm.ParameterName) = prm.ParameterValue
    Next prm
    
    ' Ejecutar consulta parametrizada
    Set rs = qdf.OpenRecordset()
    
    ' Mapear recordset a objeto de dominio
    If Not rs.EOF Then
        Set mapeo = New EMapeo
        mapeo.IdMapeo = rs!IdMapeo.value
        mapeo.NombrePlantilla = rs!NombrePlantilla.value & ""
        mapeo.NombreCampoTabla = rs!NombreCampoTabla.value & ""
        mapeo.ValorAsociado = rs!ValorAsociado.value & ""
        mapeo.NombreCampoWord = rs!NombreCampoWord.value & ""
    End If
    
    ' Limpiar recursos
    rs.Close
    Set rs = Nothing
    Set qdf = Nothing
    db.Close
    Set db = Nothing
    
    ' Devolver el objeto mapeado
    Set IMapeoRepository_GetMapeoPorTipo = mapeo
    
    Exit Function
ErrorHandler:
    ' Limpiar recursos en caso de error
    If Not rs Is Nothing Then rs.Close
    If Not qdf Is Nothing Then Set qdf = Nothing
    If Not db Is Nothing Then
        db.Close
        Set db = Nothing
    End If
    m_ErrorHandler.LogError Err.Number, Err.Description, "CMapeoRepository.GetMapeoPorTipo"
    Set IMapeoRepository_GetMapeoPorTipo = Nothing
End Function

' ============================================================================
' M?TODOS P?BLICOS DE CONVENIENCIA
' ============================================================================
' Estos m?todos exponen p?blicamente la funcionalidad de la interfaz
' para facilitar las pruebas y el uso directo de la clase concreta

' M?todo p?blico de conveniencia para GetMapeoPorTipo
Public Function GetMapeoPorTipo(ByVal tipoSolicitud As String) As EMapeo
    Set GetMapeoPorTipo = IMapeoRepository_GetMapeoPorTipo(tipoSolicitud)
End Function

Private Function IMapeoRepository_ObtenerMapeosPorCategoria(ByVal categoria As String) As Object
    ' Placeholder para compilaci?n
    Set IMapeoRepository_ObtenerMapeosPorCategoria = New Scripting.Dictionary
End Function

Public Function ObtenerMapeosPorCategoria(ByVal categoria As String) As Object
    Set ObtenerMapeosPorCategoria = IMapeoRepository_ObtenerMapeosPorCategoria(categoria)
End Function
