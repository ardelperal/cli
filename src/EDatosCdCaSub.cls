VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EDatosCdCaSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private m_idDatosCDCASUB As Long
Private m_idSolicitud As Long
Private m_refSuministrador As String
Private m_refSubSuministrador As String
Private m_suministradorPrincipalNombreDir As String
Private m_subSuministradorNombreDir As String
Private m_identificacionMaterial As String
Private m_numPlanoEspecificacion As String
Private m_cantidadPeriodo As Long
Private m_numSerieLote As String
Private m_descripcionImpactoNC As String
Private m_descripcionImpactoNCCont As String
Private m_refDesviacionesPrevias As String
Private m_causaNC As String
Private m_impactoCoste As String
Private m_clasificacionNC As String
Private m_afectaPrestaciones As Boolean
Private m_afectaSeguridad As Boolean
Private m_afectaFiabilidad As Boolean
Private m_afectaVidaUtil As Boolean
Private m_afectaMedioambiente As Boolean
Private m_afectaIntercambiabilidad As Boolean
Private m_afectaMantenibilidad As Boolean
Private m_afectaApariencia As Boolean
Private m_afectaOtros As Boolean
Private m_requiereModificacionContrato As Boolean
Private m_efectoFechaEntrega As String
Private m_identificacionAutoridadDiseno As String
Private m_esSubSuministradorAD As Boolean
Private m_nombreRepSubSuministrador As String
Private m_racRef As String
Private m_racCodigo As String
Private m_observacionesRAC As String
Private m_fechaFirmaRAC As Date
Private m_decisionSuministradorPrincipal As String
Private m_obsSuministradorPrincipal As String
Private m_fechaFirmaSuministradorPrincipal As Date
Private m_firmaSuministradorPrincipalNombreCargo As String
Private m_obsRACDelegador As String
Private m_fechaFirmaRACDelegador As Date

' Propiedades p?blicas
Public Property Get idDatosCDCASUB() As Long
    idDatosCDCASUB = m_idDatosCDCASUB
End Property

Public Property Let idDatosCDCASUB(ByVal value As Long)
    m_idDatosCDCASUB = value
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

Public Property Get refSubSuministrador() As String
    refSubSuministrador = m_refSubSuministrador
End Property

Public Property Let refSubSuministrador(ByVal value As String)
    m_refSubSuministrador = value
End Property

Public Property Get suministradorPrincipalNombreDir() As String
    suministradorPrincipalNombreDir = m_suministradorPrincipalNombreDir
End Property

Public Property Let suministradorPrincipalNombreDir(ByVal value As String)
    m_suministradorPrincipalNombreDir = value
End Property

Public Property Get subSuministradorNombreDir() As String
    subSuministradorNombreDir = m_subSuministradorNombreDir
End Property

Public Property Let subSuministradorNombreDir(ByVal value As String)
    m_subSuministradorNombreDir = value
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

Public Property Get cantidadPeriodo() As Long
    cantidadPeriodo = m_cantidadPeriodo
End Property

Public Property Let cantidadPeriodo(ByVal value As Long)
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

Public Property Get afectaPrestaciones() As Boolean
    afectaPrestaciones = m_afectaPrestaciones
End Property

Public Property Let afectaPrestaciones(ByVal value As Boolean)
    m_afectaPrestaciones = value
End Property

Public Property Get afectaSeguridad() As Boolean
    afectaSeguridad = m_afectaSeguridad
End Property

Public Property Let afectaSeguridad(ByVal value As Boolean)
    m_afectaSeguridad = value
End Property

Public Property Get afectaFiabilidad() As Boolean
    afectaFiabilidad = m_afectaFiabilidad
End Property

Public Property Let afectaFiabilidad(ByVal value As Boolean)
    m_afectaFiabilidad = value
End Property

Public Property Get afectaVidaUtil() As Boolean
    afectaVidaUtil = m_afectaVidaUtil
End Property

Public Property Let afectaVidaUtil(ByVal value As Boolean)
    m_afectaVidaUtil = value
End Property

Public Property Get afectaMedioambiente() As Boolean
    afectaMedioambiente = m_afectaMedioambiente
End Property

Public Property Let afectaMedioambiente(ByVal value As Boolean)
    m_afectaMedioambiente = value
End Property

Public Property Get afectaIntercambiabilidad() As Boolean
    afectaIntercambiabilidad = m_afectaIntercambiabilidad
End Property

Public Property Let afectaIntercambiabilidad(ByVal value As Boolean)
    m_afectaIntercambiabilidad = value
End Property

Public Property Get afectaMantenibilidad() As Boolean
    afectaMantenibilidad = m_afectaMantenibilidad
End Property

Public Property Let afectaMantenibilidad(ByVal value As Boolean)
    m_afectaMantenibilidad = value
End Property

Public Property Get afectaApariencia() As Boolean
    afectaApariencia = m_afectaApariencia
End Property

Public Property Let afectaApariencia(ByVal value As Boolean)
    m_afectaApariencia = value
End Property

Public Property Get afectaOtros() As Boolean
    afectaOtros = m_afectaOtros
End Property

Public Property Let afectaOtros(ByVal value As Boolean)
    m_afectaOtros = value
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

Public Property Get esSubSuministradorAD() As Boolean
    esSubSuministradorAD = m_esSubSuministradorAD
End Property

Public Property Let esSubSuministradorAD(ByVal value As Boolean)
    m_esSubSuministradorAD = value
End Property

Public Property Get nombreRepSubSuministrador() As String
    nombreRepSubSuministrador = m_nombreRepSubSuministrador
End Property

Public Property Let nombreRepSubSuministrador(ByVal value As String)
    m_nombreRepSubSuministrador = value
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

Public Property Get decisionSuministradorPrincipal() As String
    decisionSuministradorPrincipal = m_decisionSuministradorPrincipal
End Property

Public Property Let decisionSuministradorPrincipal(ByVal value As String)
    m_decisionSuministradorPrincipal = value
End Property

Public Property Get obsSuministradorPrincipal() As String
    obsSuministradorPrincipal = m_obsSuministradorPrincipal
End Property

Public Property Let obsSuministradorPrincipal(ByVal value As String)
    m_obsSuministradorPrincipal = value
End Property

Public Property Get fechaFirmaSuministradorPrincipal() As Date
    fechaFirmaSuministradorPrincipal = m_fechaFirmaSuministradorPrincipal
End Property

Public Property Let fechaFirmaSuministradorPrincipal(ByVal value As Date)
    m_fechaFirmaSuministradorPrincipal = value
End Property

Public Property Get firmaSuministradorPrincipalNombreCargo() As String
    firmaSuministradorPrincipalNombreCargo = m_firmaSuministradorPrincipalNombreCargo
End Property

Public Property Let firmaSuministradorPrincipalNombreCargo(ByVal value As String)
    m_firmaSuministradorPrincipalNombreCargo = value
End Property

Public Property Get obsRACDelegador() As String
    obsRACDelegador = m_obsRACDelegador
End Property

Public Property Let obsRACDelegador(ByVal value As String)
    m_obsRACDelegador = value
End Property

Public Property Get fechaFirmaRACDelegador() As Date
    fechaFirmaRACDelegador = m_fechaFirmaRACDelegador
End Property

Public Property Let fechaFirmaRACDelegador(ByVal value As Date)
    m_fechaFirmaRACDelegador = value
End Property
