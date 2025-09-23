VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockExpedienteService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IExpedienteService

Private m_GetById_Result As EExpediente
Private m_GetByNemo_Result As EExpediente
Private m_GetSelector_Result As Object

Private Sub Class_Initialize(): Call Reset: End Sub
Public Sub Reset()
    Set m_GetById_Result = Nothing
    Set m_GetByNemo_Result = Nothing
    Set m_GetSelector_Result = Nothing
End Sub

Public Sub ConfigureObtenerExpedientePorId(ByVal value As EExpediente)
    Set m_GetById_Result = value
End Sub
Public Sub ConfigureObtenerExpedientePorNemotecnico(ByVal value As EExpediente)
    Set m_GetByNemo_Result = value
End Sub
Public Sub ConfigureObtenerExpedientesParaSelector(ByVal value As Object)
    Set m_GetSelector_Result = value
End Sub

' --- IMPLEMENTACI?N DE LA INTERFAZ (PRIVADA) ---

Private Function IExpedienteService_ObtenerExpedientePorId(ByVal idExpediente As Long) As EExpediente
    Set IExpedienteService_ObtenerExpedientePorId = m_GetById_Result
End Function

Private Function IExpedienteService_ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As EExpediente
    Set IExpedienteService_ObtenerExpedientePorNemotecnico = m_GetByNemo_Result
End Function

Private Function IExpedienteService_ObtenerExpedientesParaSelector() As Object
    ' CORRECCI?N: El nombre del m?todo ahora coincide exactamente con la interfaz.
    Set IExpedienteService_ObtenerExpedientesParaSelector = m_GetSelector_Result
End Function
