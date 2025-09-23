VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "INotificationRepository"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Public Function EnqueueEmail(ByVal destinatarios As String, _
                           ByVal asunto As String, _
                           ByVal cuerpoHTML As String, _
                           Optional ByVal conCopia As String = "", _
                           Optional ByVal conCopiaOculta As String = "", _
                           Optional ByVal adjuntoPath As String = "") As Boolean
End Function

