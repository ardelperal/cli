VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CMockNotificationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

#If DEV_MODE Then

' CMockNotificationService.cls
' Mock del servicio de notificaciones para pruebas de integraci?n
' Simula el comportamiento de CNotificationService sin acceder a bases de datos externas


' Implementa la interfaz INotificationService
Implements INotificationService

' Variables para tracking de llamadas en tests
Public UltimaLlamada_Destinatarios As String
Public UltimaLlamada_Asunto As String
Public UltimaLlamada_CuerpoHTML As String
Public UltimaLlamada_URLAdjunto As String
Public UltimaLlamada_ConCopia As String
Public UltimaLlamada_ConCopiaOculta As String
Public NumeroLlamadas As Long
Public SimularError As Boolean
Public ValorRetorno As Boolean

' Constructor
Private Sub Class_Initialize()
    ' Inicializar valores por defecto
    Me.Reset
    SimularError = False
    ValorRetorno = True
End Sub



' M?todo para limpiar el historial de llamadas
Public Sub Reset()
    UltimaLlamada_Destinatarios = ""
    UltimaLlamada_Asunto = ""
    UltimaLlamada_CuerpoHTML = ""
    UltimaLlamada_URLAdjunto = ""
    UltimaLlamada_ConCopia = ""
    UltimaLlamada_ConCopiaOculta = ""
    NumeroLlamadas = 0
End Sub

' Implementaci?n mock del m?todo SendNotification
Private Function INotificationService_SendNotification( _
    ByVal destinatarios As String, _
    ByVal asunto As String, _
    ByVal cuerpoHTML As String, _
    Optional ByVal conCopia As String = "", _
    Optional ByVal conCopiaOculta As String = "", _
    Optional ByVal urlAdjunto As String = "" _
) As Boolean
    
    ' Incrementar contador de llamadas
    NumeroLlamadas = NumeroLlamadas + 1
    
    ' Guardar par?metros de la ?ltima llamada
    UltimaLlamada_Destinatarios = destinatarios
    UltimaLlamada_Asunto = asunto
    UltimaLlamada_CuerpoHTML = cuerpoHTML
    UltimaLlamada_URLAdjunto = urlAdjunto
    UltimaLlamada_ConCopia = conCopia
    UltimaLlamada_ConCopiaOculta = conCopiaOculta
    
    ' Simular error si est? configurado
    If SimularError Then
        Err.Raise 9999, "CMockNotificationService", "Error simulado para pruebas"
    End If
    
    ' Retornar valor configurado
    INotificationService_SendNotification = ValorRetorno
End Function

' Implementaci?n p?blica del m?todo SendNotification
Public Function SendNotification( _
    ByVal destinatarios As String, _
    ByVal asunto As String, _
    ByVal cuerpoHTML As String, _
    Optional ByVal conCopia As String = "", _
    Optional ByVal conCopiaOculta As String = "", _
    Optional ByVal urlAdjunto As String = "" _
) As Boolean
    ' Delegar a la implementaci?n de la interfaz
    SendNotification = INotificationService_SendNotification(destinatarios, asunto, cuerpoHTML, conCopia, conCopiaOculta, urlAdjunto)
End Function

' M?todo helper para verificar si se llam? con par?metros espec?ficos
Public Function FueLlamadoCon( _
    ByVal destinatarios As String, _
    ByVal asunto As String, _
    ByVal cuerpoHTML As String, _
    Optional ByVal urlAdjunto As String = "" _
) As Boolean
    
    FueLlamadoCon = (UltimaLlamada_Destinatarios = destinatarios) And _
                    (UltimaLlamada_Asunto = asunto) And _
                    (UltimaLlamada_CuerpoHTML = cuerpoHTML) And _
                    (UltimaLlamada_URLAdjunto = urlAdjunto)
End Function

#End If

