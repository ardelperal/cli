VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CExpedienteService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IExpedienteService

Private m_Repo As IExpedienteRepository
Private m_Logger As IOperationLogger
Private m_ErrorHandler As IErrorHandlerService

Public Sub Initialize(ByVal repo As IExpedienteRepository, ByVal logger As IOperationLogger, ByVal ErrorHandler As IErrorHandlerService)
    Set m_Repo = repo
    Set m_Logger = logger
    Set m_ErrorHandler = ErrorHandler
End Sub

' --- IMPLEMENTACI?N DE LA INTERFAZ (PRIVADA) ---

Private Function IExpedienteService_ObtenerExpedientePorId(ByVal idExpediente As Long) As EExpediente
    ' La implementaci?n delega la llamada al m?todo p?blico del repositorio
    Set IExpedienteService_ObtenerExpedientePorId = m_Repo.ObtenerExpedientePorId(idExpediente)
End Function

Private Function IExpedienteService_ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As EExpediente
    Set IExpedienteService_ObtenerExpedientePorNemotecnico = m_Repo.ObtenerExpedientePorNemotecnico(Nemotecnico)
End Function

Private Function IExpedienteService_ObtenerExpedientesParaSelector() As Object
    ' CORRECCI?N: El nombre del m?todo ahora coincide exactamente con la interfaz.
    ' La l?gica de "Activos" est? correctamente encapsulada en el repositorio.
    Set IExpedienteService_ObtenerExpedientesParaSelector = m_Repo.ObtenerExpedientesActivosParaSelector()
End Function

' --- M?TODOS P?BLICOS DE CONVENIENCIA (LECCI?N 24) ---

Public Function ObtenerExpedientePorId(ByVal idExpediente As Long) As EExpediente
    Set ObtenerExpedientePorId = IExpedienteService_ObtenerExpedientePorId(idExpediente)
End Function

Public Function ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As EExpediente
    Set ObtenerExpedientePorNemotecnico = IExpedienteService_ObtenerExpedientePorNemotecnico(Nemotecnico)
End Function

Public Function ObtenerExpedientesParaSelector() As Object
    Set ObtenerExpedientesParaSelector = IExpedienteService_ObtenerExpedientesParaSelector()
End Function
