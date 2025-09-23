VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EDatosPc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database

Private m_idDatosPC As Long
Private m_idSolicitud As Long
Private m_refContratoInspeccionOficial As String
Private m_refSuministrador As String
Private m_suministradorNombreDir As String
Private m_objetoContrato As String
Private m_descripcionMaterialAfectado As String
Private m_numPlanoEspecificacion As String
Private m_descripcionPropuestaCambio As String
Private m_descripcionPropuestaCambioCont As String
Private m_motivoCorregirDeficiencias As Boolean
Private m_motivoMejorarCapacidad As Boolean
Private m_motivoAumentarNacionalizacion As Boolean
Private m_motivoMejorarSeguridad As Boolean
Private m_motivoMejorarFiabilidad As Boolean
Private m_motivoMejorarCosteEficacia As Boolean
Private m_motivoOtros As Boolean
Private m_motivoOtrosDetalle As String
Private m_incidenciaCoste As String
Private m_incidenciaPlazo As String
Private m_incidenciaSeguridad As Boolean
Private m_incidenciaFiabilidad As Boolean
Private m_incidenciaMantenibilidad As Boolean
Private m_incidenciaIntercambiabilidad As Boolean
Private m_incidenciaVidaUtilAlmacen As Boolean
Private m_incidenciaFuncionamientoFuncion As Boolean
Private m_cambioAfectaMaterialEntregado As Boolean
Private m_cambioAfectaMaterialPorEntregar As Boolean
Private m_firmaOficinaTecnicaNombre As String
Private m_firmaRepSuministradorNombre As String
Private m_observacionesRACRef As String
Private m_racCodigo As String
Private m_observacionesRAC As String
Private m_fechaFirmaRAC As Date
Private m_obsAprobacionAutoridadDiseno As String
Private m_firmaAutoridadDisenoNombreCargo As String
Private m_fechaFirmaAutoridadDiseno As Date
Private m_decisionFinal As String
Private m_obsDecisionFinal As String
Private m_cargoFirmanteFinal As String
Private m_fechaFirmaDecisionFinal As Date

'******************************************************************************
' PROPIEDADES P?BLICAS
'******************************************************************************

Public Property Get idDatosPC() As Long
    idDatosPC = m_idDatosPC
End Property

Public Property Let idDatosPC(ByVal value As Long)
    m_idDatosPC = value
End Property

Public Property Get idSolicitud() As Long
    idSolicitud = m_idSolicitud
End Property

Public Property Let idSolicitud(ByVal value As Long)
    m_idSolicitud = value
End Property

Public Property Get refContratoInspeccionOficial() As String
    refContratoInspeccionOficial = m_refContratoInspeccionOficial
End Property

Public Property Let refContratoInspeccionOficial(ByVal value As String)
    m_refContratoInspeccionOficial = value
End Property

Public Property Get refSuministrador() As String
    refSuministrador = m_refSuministrador
End Property

Public Property Let refSuministrador(ByVal value As String)
    m_refSuministrador = value
End Property

Public Property Get suministradorNombreDir() As String
    suministradorNombreDir = m_suministradorNombreDir
End Property

Public Property Let suministradorNombreDir(ByVal value As String)
    m_suministradorNombreDir = value
End Property

Public Property Get objetoContrato() As String
    objetoContrato = m_objetoContrato
End Property

Public Property Let objetoContrato(ByVal value As String)
    m_objetoContrato = value
End Property

Public Property Get descripcionMaterialAfectado() As String
    descripcionMaterialAfectado = m_descripcionMaterialAfectado
End Property

Public Property Let descripcionMaterialAfectado(ByVal value As String)
    m_descripcionMaterialAfectado = value
End Property

Public Property Get numPlanoEspecificacion() As String
    numPlanoEspecificacion = m_numPlanoEspecificacion
End Property

Public Property Let numPlanoEspecificacion(ByVal value As String)
    m_numPlanoEspecificacion = value
End Property

Public Property Get descripcionPropuestaCambio() As String
    descripcionPropuestaCambio = m_descripcionPropuestaCambio
End Property

Public Property Let descripcionPropuestaCambio(ByVal value As String)
    m_descripcionPropuestaCambio = value
End Property

Public Property Get descripcionPropuestaCambioCont() As String
    descripcionPropuestaCambioCont = m_descripcionPropuestaCambioCont
End Property

Public Property Let descripcionPropuestaCambioCont(ByVal value As String)
    m_descripcionPropuestaCambioCont = value
End Property

Public Property Get motivoCorregirDeficiencias() As Boolean
    motivoCorregirDeficiencias = m_motivoCorregirDeficiencias
End Property

Public Property Let motivoCorregirDeficiencias(ByVal value As Boolean)
    m_motivoCorregirDeficiencias = value
End Property

Public Property Get motivoMejorarCapacidad() As Boolean
    motivoMejorarCapacidad = m_motivoMejorarCapacidad
End Property

Public Property Let motivoMejorarCapacidad(ByVal value As Boolean)
    m_motivoMejorarCapacidad = value
End Property

Public Property Get motivoAumentarNacionalizacion() As Boolean
    motivoAumentarNacionalizacion = m_motivoAumentarNacionalizacion
End Property

Public Property Let motivoAumentarNacionalizacion(ByVal value As Boolean)
    m_motivoAumentarNacionalizacion = value
End Property

Public Property Get motivoMejorarSeguridad() As Boolean
    motivoMejorarSeguridad = m_motivoMejorarSeguridad
End Property

Public Property Let motivoMejorarSeguridad(ByVal value As Boolean)
    m_motivoMejorarSeguridad = value
End Property

Public Property Get motivoMejorarFiabilidad() As Boolean
    motivoMejorarFiabilidad = m_motivoMejorarFiabilidad
End Property

Public Property Let motivoMejorarFiabilidad(ByVal value As Boolean)
    m_motivoMejorarFiabilidad = value
End Property

Public Property Get motivoMejorarCosteEficacia() As Boolean
    motivoMejorarCosteEficacia = m_motivoMejorarCosteEficacia
End Property

Public Property Let motivoMejorarCosteEficacia(ByVal value As Boolean)
    m_motivoMejorarCosteEficacia = value
End Property

Public Property Get motivoOtros() As Boolean
    motivoOtros = m_motivoOtros
End Property

Public Property Let motivoOtros(ByVal value As Boolean)
    m_motivoOtros = value
End Property

Public Property Get motivoOtrosDetalle() As String
    motivoOtrosDetalle = m_motivoOtrosDetalle
End Property

Public Property Let motivoOtrosDetalle(ByVal value As String)
    m_motivoOtrosDetalle = value
End Property

Public Property Get incidenciaCoste() As String
    incidenciaCoste = m_incidenciaCoste
End Property

Public Property Let incidenciaCoste(ByVal value As String)
    m_incidenciaCoste = value
End Property

Public Property Get incidenciaPlazo() As String
    incidenciaPlazo = m_incidenciaPlazo
End Property

Public Property Let incidenciaPlazo(ByVal value As String)
    m_incidenciaPlazo = value
End Property

Public Property Get incidenciaSeguridad() As Boolean
    incidenciaSeguridad = m_incidenciaSeguridad
End Property

Public Property Let incidenciaSeguridad(ByVal value As Boolean)
    m_incidenciaSeguridad = value
End Property

Public Property Get incidenciaFiabilidad() As Boolean
    incidenciaFiabilidad = m_incidenciaFiabilidad
End Property

Public Property Let incidenciaFiabilidad(ByVal value As Boolean)
    m_incidenciaFiabilidad = value
End Property

Public Property Get incidenciaMantenibilidad() As Boolean
    incidenciaMantenibilidad = m_incidenciaMantenibilidad
End Property

Public Property Let incidenciaMantenibilidad(ByVal value As Boolean)
    m_incidenciaMantenibilidad = value
End Property

Public Property Get incidenciaIntercambiabilidad() As Boolean
    incidenciaIntercambiabilidad = m_incidenciaIntercambiabilidad
End Property

Public Property Let incidenciaIntercambiabilidad(ByVal value As Boolean)
    m_incidenciaIntercambiabilidad = value
End Property

Public Property Get incidenciaVidaUtilAlmacen() As Boolean
    incidenciaVidaUtilAlmacen = m_incidenciaVidaUtilAlmacen
End Property

Public Property Let incidenciaVidaUtilAlmacen(ByVal value As Boolean)
    m_incidenciaVidaUtilAlmacen = value
End Property

Public Property Get incidenciaFuncionamientoFuncion() As Boolean
    incidenciaFuncionamientoFuncion = m_incidenciaFuncionamientoFuncion
End Property

Public Property Let incidenciaFuncionamientoFuncion(ByVal value As Boolean)
    m_incidenciaFuncionamientoFuncion = value
End Property

Public Property Get cambioAfectaMaterialEntregado() As Boolean
    cambioAfectaMaterialEntregado = m_cambioAfectaMaterialEntregado
End Property

Public Property Let cambioAfectaMaterialEntregado(ByVal value As Boolean)
    m_cambioAfectaMaterialEntregado = value
End Property

Public Property Get cambioAfectaMaterialPorEntregar() As Boolean
    cambioAfectaMaterialPorEntregar = m_cambioAfectaMaterialPorEntregar
End Property

Public Property Let cambioAfectaMaterialPorEntregar(ByVal value As Boolean)
    m_cambioAfectaMaterialPorEntregar = value
End Property

Public Property Get firmaOficinaTecnicaNombre() As String
    firmaOficinaTecnicaNombre = m_firmaOficinaTecnicaNombre
End Property

Public Property Let firmaOficinaTecnicaNombre(ByVal value As String)
    m_firmaOficinaTecnicaNombre = value
End Property

Public Property Get firmaRepSuministradorNombre() As String
    firmaRepSuministradorNombre = m_firmaRepSuministradorNombre
End Property

Public Property Let firmaRepSuministradorNombre(ByVal value As String)
    m_firmaRepSuministradorNombre = value
End Property

Public Property Get observacionesRACRef() As String
    observacionesRACRef = m_observacionesRACRef
End Property

Public Property Let observacionesRACRef(ByVal value As String)
    m_observacionesRACRef = value
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

Public Property Get obsAprobacionAutoridadDiseno() As String
    obsAprobacionAutoridadDiseno = m_obsAprobacionAutoridadDiseno
End Property

Public Property Let obsAprobacionAutoridadDiseno(ByVal value As String)
    m_obsAprobacionAutoridadDiseno = value
End Property

Public Property Get firmaAutoridadDisenoNombreCargo() As String
    firmaAutoridadDisenoNombreCargo = m_firmaAutoridadDisenoNombreCargo
End Property

Public Property Let firmaAutoridadDisenoNombreCargo(ByVal value As String)
    m_firmaAutoridadDisenoNombreCargo = value
End Property

Public Property Get fechaFirmaAutoridadDiseno() As Date
    fechaFirmaAutoridadDiseno = m_fechaFirmaAutoridadDiseno
End Property

Public Property Let fechaFirmaAutoridadDiseno(ByVal value As Date)
    m_fechaFirmaAutoridadDiseno = value
End Property

Public Property Get decisionFinal() As String
    decisionFinal = m_decisionFinal
End Property

Public Property Let decisionFinal(ByVal value As String)
    m_decisionFinal = value
End Property

Public Property Get obsDecisionFinal() As String
    obsDecisionFinal = m_obsDecisionFinal
End Property

Public Property Let obsDecisionFinal(ByVal value As String)
    m_obsDecisionFinal = value
End Property

Public Property Get cargoFirmanteFinal() As String
    cargoFirmanteFinal = m_cargoFirmanteFinal
End Property

Public Property Let cargoFirmanteFinal(ByVal value As String)
    m_cargoFirmanteFinal = value
End Property

Public Property Get fechaFirmaDecisionFinal() As Date
    fechaFirmaDecisionFinal = m_fechaFirmaDecisionFinal
End Property

Public Property Let fechaFirmaDecisionFinal(ByVal value As Date)
    m_fechaFirmaDecisionFinal = value
End Property
