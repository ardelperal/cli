# Esquema de base de datos

## Tabla: ~TMPCLP14691
- **PK:** idOperacion

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| idOperacion | Long | 4 | No |  |
| fechaHora | Double | 8 | Si |  |
| usuario | Text | 100 | Si |  |
| tipoOperacion | Text | 50 | Si |  |
| entidad | Text | 50 | Si |  |
| idEntidad | Long | 4 | No |  |
| descripcion | OLE Object | 0 | Si |  |
| resultado | Text | 20 | Si |  |
| detalles | OLE Object | 0 | No |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

## Tabla: ~TMPCLP19231
- **PK:** ID_Solicitud

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| ID_Solicitud | Long | 4 | No |  |
| ID_Expediente | Text | 255 | No |  |
| TipoSolicitud | Text | 50 | No |  |
| SubTipoSolicitud | Text | 50 | No |  |
| CodigoSolicitud | Text | 255 | No |  |
| EstadoInterno | Text | 50 | No |  |
| FechaCreacion | Double | 8 | No |  |
| UsuarioCreacion | Text | 255 | No |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

## Tabla: ~TMPCLP371151
- **PK:** idLogError

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| idLogError | Long | 4 | No |  |
| fechaHora | Double | 8 | Si |  |
| usuario | Text | 100 | No |  |
| modulo | Text | 100 | Si |  |
| procedimiento | Text | 100 | No |  |
| numeroError | Long | 4 | Si |  |
| descripcionError | OLE Object | 0 | Si |  |
| contexto | OLE Object | 0 | No |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

## Tabla: ~TMPCLP590901
- **PK:** idAdjunto

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| idAdjunto | Long | 4 | No |  |
| idSolicitud | Long | 4 | Si |  |
| nombreArchivo | Text | 255 | Si |  |
| tipoArchivo | Text | 50 | No |  |
| tamanoBytes | Long | 4 | No |  |
| fechaSubida | Double | 8 | Si |  |
| usuarioSubida | Text | 100 | Si |  |
| descripcion | OLE Object | 0 | No |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

## Tabla: ~TMPCLP95081
- **PK:** idSolicitud

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| idSolicitud | Long | 4 | No |  |
| idExpediente | Long | 4 | Si |  |
| tipoSolicitud | Text | 255 | Si |  |
| subTipoSolicitud | Text | 255 | No |  |
| codigoSolicitud | Text | 50 | Si |  |
| estadoInterno | Text | 50 | No |  |
| fechaCreacion | Double | 8 | Si |  |
| usuarioCreacion | Text | 100 | Si |  |
| fechaPaseTecnico | Double | 8 | No |  |
| fechaCompletadoTecnico | Double | 8 | No |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

## Tabla: tbAdjuntos
- **PK:** idAdjunto

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| idAdjunto | Long | 4 | No |  |
| idSolicitud | Long | 4 | Si |  |
| nombreArchivo | Text | 255 | Si |  |
| fechaSubida | Double | 8 | Si |  |
| usuarioSubida | Text | 100 | Si |  |
| descripcion | OLE Object | 0 | No |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

## Tabla: tbConfiguracion
- **PK:** idConfiguracion

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| idConfiguracion | Long | 4 | No |  |
| clave | Text | 255 | Si |  |
| valor | OLE Object | 0 | No |  |
| descripcion | Text | 255 | No |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

## Tabla: tbDatosCDCA
- **PK:** idDatosCDCA

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| idDatosCDCA | Long | 4 | No |  |
| idSolicitud | Long | 4 | Si |  |
| refSuministrador | Text | 100 | No |  |
| numContrato | Text | 100 | No |  |
| identificacionMaterial | OLE Object | 0 | No |  |
| numPlanoEspecificacion | Text | 100 | No |  |
| cantidadPeriodo | Text | 50 | No |  |
| numSerieLote | Text | 100 | No |  |
| descripcionImpactoNC | OLE Object | 0 | No |  |
| descripcionImpactoNCCont | OLE Object | 0 | No |  |
| refDesviacionesPrevias | Text | 100 | No |  |
| causaNC | OLE Object | 0 | No |  |
| impactoCoste | Text | 50 | No |  |
| clasificacionNC | Text | 50 | No |  |
| requiereModificacionContrato | Boolean | 1 | No |  |
| efectoFechaEntrega | OLE Object | 0 | No |  |
| identificacionAutoridadDiseno | Text | 100 | No |  |
| esSuministradorAD | Boolean | 1 | No |  |
| racRef | Text | 100 | No |  |
| racCodigo | Text | 50 | No |  |
| observacionesRAC | OLE Object | 0 | No |  |
| fechaFirmaRAC | Double | 8 | No |  |
| decisionFinal | Text | 50 | No |  |
| observacionesFinales | OLE Object | 0 | No |  |
| fechaFirmaDecisionFinal | Double | 8 | No |  |
| cargoFirmanteFinal | Text | 100 | No |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

## Tabla: tbDatosCDCASUB
- **PK:** idDatosCDCASUB

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| idDatosCDCASUB | Long | 4 | No |  |
| idSolicitud | Long | 4 | Si |  |
| refSuministrador | Text | 100 | No |  |
| refSubSuministrador | Text | 100 | No |  |
| suministradorPrincipalNombreDir | OLE Object | 0 | No |  |
| subSuministradorNombreDir | OLE Object | 0 | No |  |
| identificacionMaterial | OLE Object | 0 | No |  |
| numPlanoEspecificacion | Text | 100 | No |  |
| cantidadPeriodo | Text | 50 | No |  |
| numSerieLote | Text | 100 | No |  |
| descripcionImpactoNC | OLE Object | 0 | No |  |
| descripcionImpactoNCCont | OLE Object | 0 | No |  |
| refDesviacionesPrevias | Text | 100 | No |  |
| causaNC | OLE Object | 0 | No |  |
| impactoCoste | Text | 50 | No |  |
| clasificacionNC | Text | 50 | No |  |
| afectaPrestaciones | Boolean | 1 | No |  |
| afectaSeguridad | Boolean | 1 | No |  |
| afectaFiabilidad | Boolean | 1 | No |  |
| afectaVidaUtil | Boolean | 1 | No |  |
| afectaMedioambiente | Boolean | 1 | No |  |
| afectaIntercambiabilidad | Boolean | 1 | No |  |
| afectaMantenibilidad | Boolean | 1 | No |  |
| afectaApariencia | Boolean | 1 | No |  |
| afectaOtros | Boolean | 1 | No |  |
| requiereModificacionContrato | Boolean | 1 | No |  |
| efectoFechaEntrega | OLE Object | 0 | No |  |
| identificacionAutoridadDiseno | Text | 100 | No |  |
| esSubSuministradorAD | Boolean | 1 | No |  |
| nombreRepSubSuministrador | Text | 100 | No |  |
| racRef | Text | 100 | No |  |
| racCodigo | Text | 50 | No |  |
| observacionesRAC | OLE Object | 0 | No |  |
| fechaFirmaRAC | Double | 8 | No |  |
| decisionSuministradorPrincipal | Text | 50 | No |  |
| obsSuministradorPrincipal | OLE Object | 0 | No |  |
| fechaFirmaSuministradorPrincipal | Double | 8 | No |  |
| firmaSuministradorPrincipalNombreCargo | Text | 100 | No |  |
| obsRACDelegador | OLE Object | 0 | No |  |
| fechaFirmaRACDelegador | Double | 8 | No |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

## Tabla: tbDatosPC
- **PK:** idDatosPC

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| idDatosPC | Long | 4 | No |  |
| idSolicitud | Long | 4 | Si |  |
| refContratoInspeccionOficial | Text | 100 | No |  |
| refSuministrador | Text | 100 | No |  |
| suministradorNombreDir | OLE Object | 0 | No |  |
| objetoContrato | OLE Object | 0 | No |  |
| descripcionMaterialAfectado | OLE Object | 0 | No |  |
| numPlanoEspecificacion | Text | 100 | No |  |
| descripcionPropuestaCambio | OLE Object | 0 | No |  |
| descripcionPropuestaCambioCont | OLE Object | 0 | No |  |
| motivoCorregirDeficiencias | Boolean | 1 | No |  |
| motivoMejorarCapacidad | Boolean | 1 | No |  |
| motivoAumentarNacionalizacion | Boolean | 1 | No |  |
| motivoMejorarSeguridad | Boolean | 1 | No |  |
| motivoMejorarFiabilidad | Boolean | 1 | No |  |
| motivoMejorarCosteEficacia | Boolean | 1 | No |  |
| motivoOtros | Boolean | 1 | No |  |
| motivoOtrosDetalle | Text | 255 | No |  |
| incidenciaCoste | Text | 50 | No |  |
| incidenciaPlazo | Text | 50 | No |  |
| incidenciaSeguridad | Boolean | 1 | No |  |
| incidenciaFiabilidad | Boolean | 1 | No |  |
| incidenciaMantenibilidad | Boolean | 1 | No |  |
| incidenciaIntercambiabilidad | Boolean | 1 | No |  |
| incidenciaVidaUtilAlmacen | Boolean | 1 | No |  |
| incidenciaFuncionamientoFuncion | Boolean | 1 | No |  |
| cambioAfectaMaterialEntregado | Boolean | 1 | No |  |
| cambioAfectaMaterialPorEntregar | Boolean | 1 | No |  |
| firmaOficinaTecnicaNombre | Text | 100 | No |  |
| firmaRepSuministradorNombre | Text | 100 | No |  |
| observacionesRACRef | Text | 100 | No |  |
| racCodigo | Text | 50 | No |  |
| observacionesRAC | OLE Object | 0 | No |  |
| fechaFirmaRAC | Double | 8 | No |  |
| obsAprobacionAutoridadDiseno | OLE Object | 0 | No |  |
| firmaAutoridadDisenoNombreCargo | Text | 100 | No |  |
| fechaFirmaAutoridadDiseno | Double | 8 | No |  |
| decisionFinal | Text | 50 | No |  |
| obsDecisionFinal | OLE Object | 0 | No |  |
| cargoFirmanteFinal | Text | 100 | No |  |
| fechaFirmaDecisionFinal | Double | 8 | No |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

## Tabla: tbEstados
- **PK:** idEstado

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| idEstado | Long | 4 | No |  |
| nombreEstado | Text | 50 | Si |  |
| descripcion | Text | 255 | No |  |
| esEstadoInicial | Boolean | 1 | No | FALSE |
| esEstadoFinal | Boolean | 1 | No | FALSE |
| orden | Long | 4 | No |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

## Tabla: TbLocalConfig
- **PK:** ID

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| ID | Long | 4 | No |  |
| Entorno | Text | 20 | Si |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

## Tabla: tbLogCambios
- **PK:** idLogCambio

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| idLogCambio | Long | 4 | No |  |
| fechaHora | Double | 8 | Si |  |
| usuario | Text | 100 | Si |  |
| tabla | Text | 50 | Si |  |
| registro | Long | 4 | Si |  |
| campo | Text | 50 | No |  |
| valorAnterior | OLE Object | 0 | No |  |
| valorNuevo | OLE Object | 0 | No |  |
| tipoOperacion | Text | 20 | Si |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

## Tabla: tbLogErrores
- **PK:** idLogError

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| idLogError | Long | 4 | No |  |
| fechaHora | Double | 8 | Si |  |
| usuario | Text | 100 | No |  |
| modulo | Text | 100 | Si |  |
| procedimiento | Text | 100 | No |  |
| numeroError | Long | 4 | Si |  |
| descripcionError | OLE Object | 0 | Si |  |
| contexto | OLE Object | 0 | No |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

## Tabla: tbMapeoCampos
- **PK:** idMapeo

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| idMapeo | Long | 4 | No |  |
| nombrePlantilla | Text | 50 | Si |  |
| nombreCampoTabla | Text | 100 | Si |  |
| valorAsociado | Text | 100 | No |  |
| nombreCampoWord | Text | 100 | Si |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

## Tabla: tbOperacionesLog
- **PK:** idOperacion

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| idOperacion | Long | 4 | No |  |
| fechaHora | Double | 8 | Si |  |
| usuario | Text | 100 | Si |  |
| tipoOperacion | Text | 50 | Si |  |
| entidad | Text | 50 | Si |  |
| idEntidad | Long | 4 | No |  |
| descripcion | OLE Object | 0 | Si |  |
| resultado | Text | 20 | Si |  |
| detalles | OLE Object | 0 | No |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

## Tabla: tbSolicitudes
- **PK:** idSolicitud

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| idSolicitud | Long | 4 | No |  |
| idExpediente | Long | 4 | Si |  |
| tipoSolicitud | Text | 20 | Si |  |
| subTipoSolicitud | Text | 20 | No |  |
| codigoSolicitud | Text | 50 | Si |  |
| idEstadoInterno | Long | 4 | Si | 0 |
| fechaCreacion | Double | 8 | Si |  |
| usuarioCreacion | Text | 100 | Si |  |
| fechaPaseTecnico | Double | 8 | No |  |
| fechaCompletadoTecnico | Double | 8 | No |  |
| fechaModificacion | Double | 8 | No |  |
| usuarioModificacion | Text | 100 | No |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

## Tabla: tbTransiciones
- **PK:** idTransicion

| Campo | Tipo | Tamano | Requerido | Defecto |
|---|---|---:|:---:|---|
| idTransicion | Long | 4 | No |  |
| idEstadoOrigen | Long | 4 | Si |  |
| idEstadoDestino | Long | 4 | Si |  |
| rolRequerido | Text | 50 | Si |  |

**FK (salientes):**
- (ninguna)

**Referenciada por (entrantes):**
- (ninguna)

