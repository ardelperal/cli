VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EDatosCdCa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private m_idDatosCDCA As Long
Private m_idSolicitud As Long
Private m_refSuministrador As String
Private m_numContrato As String
Private m_identificacionMaterial As String
Private m_numPlanoEspecificacion As String
Private m_cantidadPeriodo As String
Private m_numSerieLote As String
Private m_descripcionImpactoNC As String
Private m_descripcionImpactoNCCont As String
Private m_refDesviacionesPrevias As String
Private m_causaNC As String
Private m_impactoCoste As String
Private m_clasificacionNC As String
Private m_requiereModificacionContrato As Boolean
Private m_efectoFechaEntrega As String
Private m_identificacionAutoridadDiseno As String
Private m_esSuministradorAD As Boolean
Private m_racRef As String
Private m_racCodigo As String
Private m_observacionesRAC As String
Private m_fechaFirmaRAC As Date
Private m_decisionFinal As String
Private m_observacionesFinales As String
Private m_fechaFirmaDecisionFinal As Date
Private m_cargoFirmanteFinal As String

'******************************************************************************
' PROPIEDADES P?BLICAS
'******************************************************************************

Public Property Get idDatosCDCA() As Long
    idDatosCDCA = m_idDatosCDCA
End Property

Public Property Let idDatosCDCA(ByVal value As Long)
    m_idDatosCDCA = value
End Property

Public Property Get idSolicitud() As Long
    idSolicitud = m_idSolicitud
End Property

Public Property Let idSolicitud(ByVal value As Long)
    m_idSolicitud = value
End Property

Public Property Get refSuministrador() As String
    refSuministrador = m_refSuministrador
End Property

Public Property Let refSuministrador(ByVal value As String)
    m_refSuministrador = value
End Property

Public Property Get numContrato() As String
    numContrato = m_numContrato
End Property

Public Property Let numContrato(ByVal value As String)
    m_numContrato = value
End Property

Public Property Get identificacionMaterial() As String
    identificacionMaterial = m_identificacionMaterial
End Property

Public Property Let identificacionMaterial(ByVal value As String)
    m_identificacionMaterial = value
End Property

Public Property Get numPlanoEspecificacion() As String
    numPlanoEspecificacion = m_numPlanoEspecificacion
End Property

Public Property Let numPlanoEspecificacion(ByVal value As String)
    m_numPlanoEspecificacion = value
End Property

Public Property Get cantidadPeriodo() As String
    cantidadPeriodo = m_cantidadPeriodo
End Property

Public Property Let cantidadPeriodo(ByVal value As String)
    m_cantidadPeriodo = value
End Property

Public Property Get numSerieLote() As String
    numSerieLote = m_numSerieLote
End Property

Public Property Let numSerieLote(ByVal value As String)
    m_numSerieLote = value
End Property

Public Property Get descripcionImpactoNC() As String
    descripcionImpactoNC = m_descripcionImpactoNC
End Property

Public Property Let descripcionImpactoNC(ByVal value As String)
    m_descripcionImpactoNC = value
End Property

Public Property Get descripcionImpactoNCCont() As String
    descripcionImpactoNCCont = m_descripcionImpactoNCCont
End Property

Public Property Let descripcionImpactoNCCont(ByVal value As String)
    m_descripcionImpactoNCCont = value
End Property

Public Property Get refDesviacionesPrevias() As String
    refDesviacionesPrevias = m_refDesviacionesPrevias
End Property

Public Property Let refDesviacionesPrevias(ByVal value As String)
    m_refDesviacionesPrevias = value
End Property

Public Property Get causaNC() As String
    causaNC = m_causaNC
End Property

Public Property Let causaNC(ByVal value As String)
    m_causaNC = value
End Property

Public Property Get impactoCoste() As String
    impactoCoste = m_impactoCoste
End Property

Public Property Let impactoCoste(ByVal value As String)
    m_impactoCoste = value
End Property

Public Property Get clasificacionNC() As String
    clasificacionNC = m_clasificacionNC
End Property

Public Property Let clasificacionNC(ByVal value As String)
    m_clasificacionNC = value
End Property

Public Property Get requiereModificacionContrato() As Boolean
    requiereModificacionContrato = m_requiereModificacionContrato
End Property

Public Property Let requiereModificacionContrato(ByVal value As Boolean)
    m_requiereModificacionContrato = value
End Property

Public Property Get efectoFechaEntrega() As String
    efectoFechaEntrega = m_efectoFechaEntrega
End Property

Public Property Let efectoFechaEntrega(ByVal value As String)
    m_efectoFechaEntrega = value
End Property

Public Property Get identificacionAutoridadDiseno() As String
    identificacionAutoridadDiseno = m_identificacionAutoridadDiseno
End Property

Public Property Let identificacionAutoridadDiseno(ByVal value As String)
    m_identificacionAutoridadDiseno = value
End Property

Public Property Get esSuministradorAD() As Boolean
    esSuministradorAD = m_esSuministradorAD
End Property

Public Property Let esSuministradorAD(ByVal value As Boolean)
    m_esSuministradorAD = value
End Property

Public Property Get racRef() As String
    racRef = m_racRef
End Property

Public Property Let racRef(ByVal value As String)
    m_racRef = value
End Property

Public Property Get racCodigo() As String
    racCodigo = m_racCodigo
End Property

Public Property Let racCodigo(ByVal value As String)
    m_racCodigo = value
End Property

Public Property Get observacionesRAC() As String
    observacionesRAC = m_observacionesRAC
End Property

Public Property Let observacionesRAC(ByVal value As String)
    m_observacionesRAC = value
End Property

Public Property Get fechaFirmaRAC() As Date
    fechaFirmaRAC = m_fechaFirmaRAC
End Property

Public Property Let fechaFirmaRAC(ByVal value As Date)
    m_fechaFirmaRAC = value
End Property

Public Property Get decisionFinal() As String
    decisionFinal = m_decisionFinal
End Property

Public Property Let decisionFinal(ByVal value As String)
    m_decisionFinal = value
End Property

Public Property Get observacionesFinales() As String
    observacionesFinales = m_observacionesFinales
End Property

Public Property Let observacionesFinales(ByVal value As String)
    m_observacionesFinales = value
End Property

Public Property Get fechaFirmaDecisionFinal() As Date
    fechaFirmaDecisionFinal = m_fechaFirmaDecisionFinal
End Property

Public Property Let fechaFirmaDecisionFinal(ByVal value As Date)
    m_fechaFirmaDecisionFinal = value
End Property

Public Property Get cargoFirmanteFinal() As String
    cargoFirmanteFinal = m_cargoFirmanteFinal
End Property

Public Property Let cargoFirmanteFinal(ByVal value As String)
    m_cargoFirmanteFinal = value
End Property
