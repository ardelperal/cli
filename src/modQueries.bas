Attribute VB_Name = "modQueries"
Option Compare Database
Option Explicit




'''
' Módulo Central de Consultas SQL - CONDOR
' Principio: Centralización de consultas para mantenibilidad y seguridad
'''

' ============================================================================
' CONSULTAS DE AUTENTICACIÓN (Lanzadera)
' ============================================================================

Public Const GET_AUTH_DATA_BY_EMAIL As String = _
    "SELECT U.EsAdministrador, P.EsUsuarioAdministrador, P.EsUsuarioCalidad, P.EsUsuarioTecnico " & _
    "FROM TbUsuariosAplicaciones AS U LEFT JOIN TbUsuariosAplicacionesPermisos AS P " & _
    "ON U.CorreoUsuario = P.CorreoUsuario AND P.IDAplicacion=[pIdAplicacion] " & _
    "WHERE U.CorreoUsuario=[pEmail];"

' ============================================================================
' CONSULTAS DE WORKFLOW
' ============================================================================

Public Const IS_VALID_TRANSITION As String = _
    "SELECT COUNT(idTransicion) AS TransitionCount " & _
    "FROM tbTransiciones " & _
    "WHERE idEstadoOrigen = [pIdEstadoOrigen] AND idEstadoDestino = [pIdEstadoDestino] AND rolRequerido = [pRolRequerido];"

Public Const GET_NEXT_STATES As String = _
    "SELECT ED.idEstado, ED.nombreEstado " & _
    "FROM tbTransiciones T " & _
    "INNER JOIN tbEstados ED ON T.idEstadoDestino = ED.idEstado " & _
    "WHERE T.idEstadoOrigen = [pIdEstadoActual] AND T.rolRequerido = [pUsuarioRol];"

' ============================================================================
' CONSULTAS DE SOLICITUDES
' ============================================================================

Public Const GET_SOLICITUD_BY_ID As String = _
    "SELECT * FROM tbSolicitudes WHERE idSolicitud = [pIdSolicitud];"

Public Const INSERT_SOLICITUD As String = _
    "INSERT INTO tbSolicitudes (idExpediente, tipoSolicitud, subTipoSolicitud, " & _
    "codigoSolicitud, idEstadoInterno, fechaCreacion, usuarioCreacion) " & _
    "VALUES ([pIdExpediente], [pTipoSolicitud], [pSubTipoSolicitud], " & _
    "[pCodigoSolicitud], [pIdEstadoInterno], Now(), [pUsuarioCreacion]);"

Public Const UPDATE_SOLICITUD As String = _
    "UPDATE tbSolicitudes SET idExpediente=[pIdExpediente], tipoSolicitud=[pTipoSolicitud], " & _
    "subTipoSolicitud=[pSubTipoSolicitud], idEstadoInterno=[pIdEstadoInterno], " & _
    "fechaModificacion=Now(), usuarioModificacion=[pUsuarioModificacion] " & _
    "WHERE idSolicitud=[pIdSolicitud];"

Public Const GET_DATOS_PC_BY_SOLICITUD As String = _
    "SELECT * FROM tbDatosPC WHERE idSolicitud = [pIdSolicitud];"

Public Const GET_DATOS_CDCA_BY_SOLICITUD As String = _
    "SELECT * FROM tbDatosCDCA WHERE idSolicitud = [pIdSolicitud];"

Public Const GET_DATOS_CDCASUB_BY_SOLICITUD As String = _
    "SELECT * FROM tbDatosCDCASUB WHERE idSolicitud = [pIdSolicitud];"

Public Const GET_LAST_INSERT_ID As String = "SELECT @@IDENTITY;"

' ============================================================================
' CONSULTAS DE MAPEO
' ============================================================================

Public Const GET_MAPEO_POR_TIPO As String = _
    "SELECT nombreCampoTabla, nombreCampoWord, valorAsociado " & _
    "FROM tbMapeoCampos " & _
    "WHERE nombrePlantilla = [pNombrePlantilla];"

' ============================================================================
' CONSULTAS DE OPERACIONES Y LOGGING
' ============================================================================

Public Const INSERT_OPERATION_LOG As String = _
        "PARAMETERS pUsuario TEXT(255), pTipoOperacion TEXT(255), pEntidad TEXT(100), pIdEntidad LONG, pDescripcion TEXT, pResultado TEXT(50), pDetalles TEXT; " & _
        "INSERT INTO tbOperacionesLog (fechaHora, usuario, tipoOperacion, entidad, idEntidad, descripcion, resultado, detalles) " & _
        "VALUES (Now(), [pUsuario], [pTipoOperacion], [pEntidad], [pIdEntidad], [pDescripcion], [pResultado], [pDetalles]);"


