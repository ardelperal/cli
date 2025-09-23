VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CWorkflowService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Implements IWorkflowService

Private m_Repo As IWorkflowRepository
Private m_Logger As IOperationLogger
Private m_ErrorHandler As IErrorHandlerService

Public Sub Initialize(ByVal repo As IWorkflowRepository, ByVal logger As IOperationLogger, ByVal ErrorHandler As IErrorHandlerService)
    Set m_Repo = repo
    Set m_Logger = logger
    Set m_ErrorHandler = ErrorHandler
End Sub

Private Function IWorkflowService_ValidateTransition(ByVal SolicitudID As Long, ByVal estadoOrigen As String, ByVal estadoDestino As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim isValid As Boolean
    
    ' REGLA 1: (M?XIMA PRIORIDAD) Ninguna transici?n puede salir de un estado final.
    If IWorkflowService_IsEstadoFinal(estadoOrigen) Then
        isValid = False
    ' REGLA 2: (PRIORIDAD MEDIA) Administrador y Calidad pueden realizar cualquier transici?n NO-FINAL.
    ElseIf usuarioRol = "Administrador" Or usuarioRol = "Calidad" Then
        isValid = True
    ' REGLA 3: (PRIORIDAD BAJA) Para otros roles, consultar al repositorio.
    Else
        isValid = m_Repo.IsValidTransition(tipoSolicitud, estadoOrigen, estadoDestino, usuarioRol)
    End If
    
    ' Registrar la operaci?n de validaci?n
    Dim logEntry As EOperationLog
    Set logEntry = New EOperationLog
    With logEntry
        .tipoOperacion = "VALIDATE_TRANSITION"
        .idEntidadAfectada = SolicitudID
        .usuario = "" ' No disponible en este contexto, dejar vacio
        If isValid Then
            .resultado = "Success"
        Else
            .resultado = "Failure"
        End If
        .descripcion = "Validacion de transicion: " & estadoOrigen & " -> " & estadoDestino & " para rol " & usuarioRol
    End With
    m_Logger.LogOperation logEntry
    
    IWorkflowService_ValidateTransition = isValid
    
Cleanup:
    Exit Function
    
ErrorHandler:
    m_ErrorHandler.LogError Err.Number, "Error en CWorkflowService.ValidateTransition: " & Err.Description, "CWorkflowService"
    IWorkflowService_ValidateTransition = False
    GoTo Cleanup
End Function

Private Function IWorkflowService_GetNextStates(ByVal estadoActual As String, ByVal tipoSolicitud As String, ByVal usuarioRol As String) As Object
    ' L?gica mejorada para obtener siguientes estados en el flujo de 7 estados
    On Error GoTo ErrorHandler
    
    Dim dict As New Scripting.Dictionary
    
    ' Verificar si es estado final (Aprobada no tiene transiciones de salida)
    If estadoActual = "Aprobada" Then
        Set IWorkflowService_GetNextStates = dict ' Retornar diccionario vac?o
        Exit Function
    End If
    
    ' Convertir estadoActual String a Long para el repositorio
    Dim IdEstado As Long
    
    ' Mapeo de nombres de estados a IDs para el nuevo flujo
    Select Case estadoActual
        Case "Registrado": IdEstado = 1
        Case "Desarrollo": IdEstado = 2
        Case "Modificaci?n": IdEstado = 3
        Case "Validaci?n": IdEstado = 4
        Case "Revisi?n": IdEstado = 5
        Case "Formalizaci?n": IdEstado = 6
        Case "Aprobada": IdEstado = 7
        Case Else
            ' Estado no reconocido
            m_ErrorHandler.LogError 0, "Estado no reconocido: " & estadoActual, "CWorkflowService"
            Set IWorkflowService_GetNextStates = dict
            Exit Function
    End Select
    
    Set dict = m_Repo.GetNextStates(IdEstado, usuarioRol)
    Set IWorkflowService_GetNextStates = dict
    Exit Function
    
ErrorHandler:
    m_ErrorHandler.LogError Err.Number, "Error en CWorkflowService.GetNextStates: " & Err.Description, "CWorkflowService"
    Set IWorkflowService_GetNextStates = New Scripting.Dictionary
End Function

Private Function IWorkflowService_IsEstadoFinal(ByVal estadoActual As String) As Boolean
    ' Implementaci?n para verificar si un estado es final en el nuevo flujo
    IWorkflowService_IsEstadoFinal = (estadoActual = "Aprobada")
End Function
