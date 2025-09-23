VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IOperationLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Sub LogOperation(ByVal logEntry As EOperationLog)
End Sub

Public Sub LogSolicitudOperation(ByVal logEntry As EOperationLog, ByVal solicitud As ESolicitud, ByVal userId As String)
End Sub
