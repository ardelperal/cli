VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockMapeoRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IMapeoRepository

' --- Variable de estado para configurar el comportamiento ---
Private m_GetMapeoPorTipo_Result As EMapeo
Private m_ObtenerMapeosPorCategoria_Result As Object

' Constructor
Private Sub Class_Initialize()
    Reset
End Sub

' --- M?todo p?blico para configurar el mock desde los tests ---
Public Sub ConfigureGetMapeoPorTipo(ByVal mapeo As EMapeo)
    Set m_GetMapeoPorTipo_Result = mapeo
End Sub

Public Sub ConfigureObtenerMapeosPorCategoria(ByVal mapeos As Object)
    Set m_ObtenerMapeosPorCategoria_Result = mapeos
End Sub

' --- Resetea el estado del mock para aislar los tests ---
Public Sub Reset()
    Set m_GetMapeoPorTipo_Result = Nothing
    Set m_ObtenerMapeosPorCategoria_Result = Nothing
End Sub

' --- Implementaci?n de la Interfaz ---
Private Function IMapeoRepository_GetMapeoPorTipo(ByVal tipoSolicitud As String) As EMapeo
    ' Devuelve el mapeo configurado por el test
    Set IMapeoRepository_GetMapeoPorTipo = m_GetMapeoPorTipo_Result
End Function

' ============================================================================
' M?TODOS P?BLICOS DE CONVENIENCIA
' ============================================================================

' M?todo p?blico de conveniencia para GetMapeoPorTipo
Public Function GetMapeoPorTipo(ByVal tipoSolicitud As String) As EMapeo
    Set GetMapeoPorTipo = IMapeoRepository_GetMapeoPorTipo(tipoSolicitud)
End Function

Private Function IMapeoRepository_ObtenerMapeosPorCategoria(ByVal categoria As String) As Object
    Set IMapeoRepository_ObtenerMapeosPorCategoria = m_ObtenerMapeosPorCategoria_Result
End Function

Public Function ObtenerMapeosPorCategoria(ByVal categoria As String) As Object
    Set ObtenerMapeosPorCategoria = IMapeoRepository_ObtenerMapeosPorCategoria(categoria)
End Function
