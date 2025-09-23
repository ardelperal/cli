VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockWorkflowRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IWorkflowRepository

Private m_IsValid_Result As Boolean
Private m_GetNext_Result As Object

Private Sub Class_Initialize(): Call Reset: End Sub

Public Sub Reset()
    m_IsValid_Result = False
    Set m_GetNext_Result = Nothing
End Sub

Public Sub ConfigureIsValidTransition(ByVal value As Boolean)
    m_IsValid_Result = value
End Sub

Public Sub ConfigureGetNextStates(ByVal value As Object)
    Set m_GetNext_Result = value
End Sub

Private Function IWorkflowRepository_IsValidTransition(ByVal tipoSolicitud As String, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal usuarioRol As String) As Boolean
    IWorkflowRepository_IsValidTransition = m_IsValid_Result
End Function

Private Function IWorkflowRepository_GetNextStates(ByVal idEstadoActual As Long, ByVal usuarioRol As String) As Object
    Set IWorkflowRepository_GetNextStates = m_GetNext_Result
End Function
