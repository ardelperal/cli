VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockExpedienteRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Implements IExpedienteRepository

' --- Variables para Configuraci?n de Retorno ---
Private m_GetById_Result As EExpediente
Private m_GetByNemo_Result As EExpediente
Private m_GetSelector_Result As Object

' --- Variables para Verificaci?n de Llamadas (Spy Properties) ---
Private m_ObtenerPorId_WasCalled As Boolean
Private m_ObtenerPorId_LastId As Long
Private m_ObtenerPorNemotecnico_WasCalled As Boolean
Private m_ObtenerPorNemotecnico_LastNemo As String
Private m_ObtenerSelector_WasCalled As Boolean

' --- Inicializaci?n y Limpieza ---

Private Sub Class_Initialize()
    Call Reset
End Sub

Public Sub Reset()
    ' Limpiar resultados configurados
    Set m_GetById_Result = Nothing
    Set m_GetByNemo_Result = Nothing
    Set m_GetSelector_Result = Nothing
    
    ' Limpiar propiedades de verificaci?n
    m_ObtenerPorId_WasCalled = False
    m_ObtenerPorId_LastId = 0
    m_ObtenerPorNemotecnico_WasCalled = False
    m_ObtenerPorNemotecnico_LastNemo = ""
    m_ObtenerSelector_WasCalled = False
End Sub

' --- Propiedades P?blicas de Verificaci?n (Solo Lectura) ---

Public Property Get ObtenerExpedientePorId_WasCalled() As Boolean
    ObtenerExpedientePorId_WasCalled = m_ObtenerPorId_WasCalled
End Property

Public Property Get ObtenerExpedientePorId_LastId() As Long
    ObtenerExpedientePorId_LastId = m_ObtenerPorId_LastId
End Property

Public Property Get ObtenerExpedientePorNemotecnico_WasCalled() As Boolean
    ObtenerExpedientePorNemotecnico_WasCalled = m_ObtenerPorNemotecnico_WasCalled
End Property

Public Property Get ObtenerExpedientePorNemotecnico_LastNemo() As String
    ObtenerExpedientePorNemotecnico_LastNemo = m_ObtenerPorNemotecnico_LastNemo
End Property

Public Property Get ObtenerExpedientesActivosParaSelector_WasCalled() As Boolean
    ObtenerExpedientesActivosParaSelector_WasCalled = m_ObtenerSelector_WasCalled
End Property

' --- M?todos de Configuraci?n ---

Public Sub ConfigureObtenerExpedientePorId(ByVal value As EExpediente)
    Set m_GetById_Result = value
End Sub

Public Sub ConfigureObtenerExpedientePorNemotecnico(ByVal value As EExpediente)
    Set m_GetByNemo_Result = value
End Sub

Public Sub ConfigureObtenerExpedientesActivosParaSelector(ByVal value As Object)
    Set m_GetSelector_Result = value
End Sub

' --- Implementaci?n de la Interfaz ---

Private Function IExpedienteRepository_ObtenerExpedientePorId(ByVal idExpediente As Long) As EExpediente
    ' Registrar la llamada
    m_ObtenerPorId_WasCalled = True
    m_ObtenerPorId_LastId = idExpediente
    ' Devolver el resultado configurado
    Set IExpedienteRepository_ObtenerExpedientePorId = m_GetById_Result
End Function

Private Function IExpedienteRepository_ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As EExpediente
    ' Registrar la llamada
    m_ObtenerPorNemotecnico_WasCalled = True
    m_ObtenerPorNemotecnico_LastNemo = Nemotecnico
    ' Devolver el resultado configurado
    Set IExpedienteRepository_ObtenerExpedientePorNemotecnico = m_GetByNemo_Result
End Function

Private Function IExpedienteRepository_ObtenerExpedientesActivosParaSelector() As Object
    ' Registrar la llamada
    m_ObtenerSelector_WasCalled = True
    ' Devolver el resultado configurado
    Set IExpedienteRepository_ObtenerExpedientesActivosParaSelector = m_GetSelector_Result
End Function
