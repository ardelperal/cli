VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IExpedienteRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Function ObtenerExpedientePorId(ByVal idExpediente As Long) As EExpediente
End Function

Public Function ObtenerExpedientePorNemotecnico(ByVal Nemotecnico As String) As EExpediente
End Function

Public Function ObtenerExpedientesActivosParaSelector() As Object
End Function
