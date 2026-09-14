<!--#include virtual="/dal/bloqueos_dal.asp"-->
<%
' El "en curso" de una solicitud se determina igual que en el legacy confirmado
' (estaIMEIenSolicitudenCurso): cualquier estado que NO sea uno de los 3 terminales
' bloquea una nueva solicitud sobre el mismo IMEI. Los 3 terminales se resuelven
' siempre por catalogo (nunca por literal numerico, a diferencia del pseudocodigo
' original que hablaba de codigos 6/7/8 fijos).
Function ImeiEnCurso(conn, imei)
    Dim rs, codRechazada, codCancelada, codActualizada
    codRechazada = CodigoEstadoSolicitud(conn, "Rechazada")
    codCancelada = CodigoEstadoSolicitud(conn, "Cancelada")
    codActualizada = CodigoEstadoSolicitud(conn, "Actualizada")
    Set rs = EjecutarConsulta(conn, _
        "SELECT SOL.COD_INDICE_SOL FROM SOLICITUD SOL, HISTORICO_ESTADO_SOLICITUD HIST " & _
        "WHERE SOL.IMEI LIKE ? AND SOL.COD_INDICE_SOL = HIST.COD_INDICE_SOL " & _
        "AND HIST.COD_ESTADO NOT IN (?, ?, ?) AND HIST.FECHA_FIN IS NULL", _
        Array(Left(imei, 14) & "%", codRechazada, codCancelada, codActualizada))
    ImeiEnCurso = Not rs.EOF
    rs.Close
End Function

' Si el IMEI ya esta en listaNegra no es necesario crear una nueva solicitud
Function ImeiEnListaNegra(conn, imei)
    Dim rs
    Set rs = EjecutarConsulta(conn, _
        "SELECT 1 FROM LISTA_NEGRA WHERE IMEI = ?", _
        Array(imei))
    ImeiEnListaNegra = Not rs.EOF
    rs.Close
End Function


' Para una Exclusion, busca la solicitud de Inclusion original sobre el mismo IMEI
' (resuelta siempre por CLAVE de catalogo, nunca con el literal hardcodeado que tenia
' el legacy en un punto de su codigo). Devuelve un Dictionary con los datos del cliente
' de esa solicitud padre, o Nothing si no existe ninguna.
Function BuscarSolicitudPadrePorImei(conn, imei)
    Dim rs, p, codInclusion
    codInclusion = CodigoTipoSolicitudPorClave(conn, "I")
    Set rs = EjecutarConsulta(conn, _
        "SELECT S.COD_INDICE_SOL, C.NOMBRE, C.PRIMER_APELLIDO, C.SEGUNDO_APELLIDO, C.NUM_IDENTIFICACION " & _
        "FROM SOLICITUD S INNER JOIN CLIENTE C ON S.COD_CLIENTE = C.COD_CLIENTE " & _
        "WHERE S.IMEI LIKE ? AND S.COD_TIPO_SOLICITUD = ? ORDER BY S.FECHA_SOLICITUD DESC", _
        Array(Left(imei, 14) & "%", codInclusion))
    If rs.EOF Then
        Set BuscarSolicitudPadrePorImei = Nothing
    Else
        Set p = Server.CreateObject("Scripting.Dictionary")
        p("NombreCliente")   = rs("NOMBRE") & ""
        p("Apellido1Cliente") = rs("PRIMER_APELLIDO") & ""
        p("Apellido2Cliente") = rs("SEGUNDO_APELLIDO") & ""
        p("IdentCliente")    = rs("NUM_IDENTIFICACION") & ""
        Set BuscarSolicitudPadrePorImei = p
    End If
    rs.Close
End Function

' Se llama siempre desde dentro de la transaccion de CrearSolicitudCompleta - usa las
' variantes Tx (ver nota en include/data/conexion.asp sobre por que).
Function ObtenerOCrearCliente(conn, numIdentificacion, nombre, apellido1, apellido2)
    Dim rs, codCliente
    Set rs = EjecutarConsultaTx(conn, "SELECT COD_CLIENTE FROM CLIENTE WHERE NUM_IDENTIFICACION = ?", Array(numIdentificacion))
    If Not rs.EOF Then
        codCliente = CLng(rs("COD_CLIENTE"))
        rs.Close
    Else
        rs.Close
        Call EjecutarNonQueryTx(conn, _
            "INSERT INTO CLIENTE (NOMBRE, PRIMER_APELLIDO, SEGUNDO_APELLIDO, NUM_IDENTIFICACION) VALUES (?, ?, ?, ?)", _
            Array(nombre, apellido1, apellido2, numIdentificacion))
        codCliente = ObtenerUltimoId(conn)
    End If
    ObtenerOCrearCliente = codCliente
End Function

' Codigo de negocio nuevo para ToperBTR: <CLAVE_TIPO><YYYYMMDD><id con padding a 6 digitos>.
' No es una replica del formato exacto del legacy (visto en una unica captura de pantalla,
' sin el algoritmo de construccion confirmado) - se documenta aqui como un formato NUEVO,
' propio, en vez de fingir una precision que no tenemos sobre el original.
Function GenerarCodigoSolicitud(clave, codIndiceSol)
    Dim fecha
    fecha = CStr(Year(Now())) & Right("0" & Month(Now()), 2) & Right("0" & Day(Now()), 2)
    'GenerarCodigoSolicitud = UCase(clave) & fecha & Right("000000" & CStr(codIndiceSol), 6)
    GenerarCodigoSolicitud = "Y" & _
                             fecha & _
                             Right("00000" & CStr(codIndiceSol), 5) & _
                             UCase(clave)
	
End Function

' Alta completa: cliente + solicitud + historico de estado inicial, y si es Exclusion,
' la exclusion del bloqueo en LISTA_NEGRA - todo en UNA sola transaccion (igual que el
' legacy, que si envolvia bien esta operacion en insertarSolicitudWeb).
' tipoClave: "I" (Inclusion) o "E" (Exclusion).
Function CrearSolicitudCompleta(conn, tipoClave, imei, msisdn, numDenuncia, fechaRobo, fechaDenuncia, identCliente, nombreCliente, apellido1Cliente, apellido2Cliente, codIndiceUsuario, codOperador)
    Dim resultado
    Set resultado = Server.CreateObject("Scripting.Dictionary")
    resultado("Exito") = False
    resultado("Mensaje") = MensajeErrorGenerico()
    resultado("CodSolicitud") = ""

    Dim codTipoSolicitud, codEstadoInicial, codFicheroGenerico, placeholderCod
    codTipoSolicitud = CodigoTipoSolicitudPorClave(conn, tipoClave)
    codEstadoInicial = CodigoEstadoSolicitud(conn, "Valida")
    codFicheroGenerico = ValorParametroEntero(conn, "COD_FICHERO_GENERICO", 0)
    placeholderCod = Left("TMP" & GenerarTokenAleatorio(6), 15)

    Dim huboError, codCliente, codIndiceSol, codSolicitudFinal
    huboError = False

    conn.BeginTrans
    On Error Resume Next

    codCliente = ObtenerOCrearCliente(conn, identCliente, nombreCliente, apellido1Cliente, apellido2Cliente)
    If Err.Number <> 0 Then huboError = True

    If Not huboError Then
        Call EjecutarNonQueryTx(conn, _
            "INSERT INTO SOLICITUD (COD_SOLICITUD, COD_TIPO_SOLICITUD, COD_INDICE_USUARIO, MSISDN, IMEI, NUM_DENUNCIA, FECHA_DENUNCIA, FECHA_ROBO, FECHA_SOLICITUD, FECHA_CREACION, COD_CLIENTE, COD_FICHERO, OPERADOR) " & _
            "VALUES (?, ?, ?, ?, ?, ?, ?, ?, NOW(), NOW(), ?, ?, ?)", _
            Array(placeholderCod, codTipoSolicitud, codIndiceUsuario, msisdn, imei, numDenuncia, fechaDenuncia, fechaRobo, codCliente, codFicheroGenerico, codOperador))
        If Err.Number <> 0 Then huboError = True
    End If

    If Not huboError Then
        codIndiceSol = ObtenerUltimoId(conn)
        codSolicitudFinal = GenerarCodigoSolicitud(tipoClave, codIndiceSol)
        Call EjecutarNonQueryTx(conn, "UPDATE SOLICITUD SET COD_SOLICITUD = ? WHERE COD_INDICE_SOL = ?", Array(codSolicitudFinal, codIndiceSol))
        If Err.Number <> 0 Then huboError = True
    End If

    If Not huboError Then
        Call EjecutarNonQueryTx(conn, "INSERT INTO HISTORICO_ESTADO_SOLICITUD (COD_INDICE_SOL, COD_ESTADO, FECHA_INICIO) VALUES (?, ?, NOW())", Array(codIndiceSol, codEstadoInicial))
        If Err.Number <> 0 Then huboError = True
    End If

    If Not huboError And tipoClave = "E" Then
        Call ExcluirIMEI(conn, imei, "BLOQUEO GLOBAL")
        If Err.Number <> 0 Then huboError = True
    End If

    If huboError Then

		resultado("Mensaje") = _
			"Err=" & Err.Number & _
			" Desc=" & Err.Description

		conn.RollbackTrans
		On Error Goto 0
		Set CrearSolicitudCompleta = resultado
		Exit Function

	End If

    conn.CommitTrans
    On Error Goto 0

    resultado("Exito") = True
    resultado("Mensaje") = ""
    resultado("CodSolicitud") = codSolicitudFinal
    resultado("CodIndiceSol") = codIndiceSol
    Set CrearSolicitudCompleta = resultado
End Function

Function BuscarSolicitudes(conn, filtroCodSolicitud, filtroImei, filtroMsisdn, filtroTipo, filtroEstado, filtroIdentCliente, filtroNombreCliente, filtroFechaDesde, filtroFechaHasta, limite, offset)
    Dim sql, listaParams(), n
    sql = "SELECT CAST(S.COD_INDICE_SOL AS CHAR) AS INDICE_SOLICITUD, S.COD_SOLICITUD, T.DESCRIPCION AS TIPO, S.IMEI, S.MSISDN, S.FECHA_CREACION, " & _
      "CONCAT_WS(' ', C.NOMBRE, C.PRIMER_APELLIDO, C.SEGUNDO_APELLIDO) AS CLIENTE, E.DESCRIPCION AS ESTADO_VIGENTE " & _
      "FROM SOLICITUD S " & _
      "INNER JOIN TIPO_SOLICITUD_REF T ON S.COD_TIPO_SOLICITUD = T.COD_TIPO_SOLICITUD " & _
      "INNER JOIN CLIENTE C ON S.COD_CLIENTE = C.COD_CLIENTE " & _
      "INNER JOIN HISTORICO_ESTADO_SOLICITUD H ON H.COD_INDICE_SOL = S.COD_INDICE_SOL " & _
      "AND H.FECHA_INICIO = ( " & _
      "SELECT MAX(H2.FECHA_INICIO) " & _
      "FROM HISTORICO_ESTADO_SOLICITUD H2 " & _
      "WHERE H2.COD_INDICE_SOL = S.COD_INDICE_SOL " & _
      ") " & _
      "INNER JOIN ESTADO_SOLICITUD_REF E ON H.COD_ESTADO = E.COD_ESTADO " & _
      "WHERE 1=1"
    ReDim listaParams(-1)
    n = 0

    If EsCadenaNoVacia(filtroCodSolicitud) Then
        sql = sql & " AND S.COD_SOLICITUD LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroCodSolicitud & "%" : n = n + 1
    End If
    If EsCadenaNoVacia(filtroImei) Then
        sql = sql & " AND S.IMEI LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroImei & "%" : n = n + 1
    End If
    If EsCadenaNoVacia(filtroMsisdn) Then
        sql = sql & " AND S.MSISDN LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroMsisdn & "%" : n = n + 1
    End If
    If EsCadenaNoVacia(filtroTipo) And filtroTipo <> "-1" And IsNumeric(filtroTipo) Then
        sql = sql & " AND S.COD_TIPO_SOLICITUD = ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CLng(filtroTipo) : n = n + 1
    End If
    If EsCadenaNoVacia(filtroEstado) And filtroEstado <> "-1" And IsNumeric(filtroEstado) Then
        sql = sql & " AND H.COD_ESTADO = ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CLng(filtroEstado) : n = n + 1
    End If
    If EsCadenaNoVacia(filtroIdentCliente) Then
        sql = sql & " AND C.NUM_IDENTIFICACION LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroIdentCliente & "%" : n = n + 1
    End If
    If EsCadenaNoVacia(filtroNombreCliente) Then
        sql = sql & " AND C.NOMBRE LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroNombreCliente & "%" : n = n + 1
    End If
    If IsDate(filtroFechaDesde) Then
        sql = sql & " AND S.FECHA_CREACION >= ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CDate(filtroFechaDesde) : n = n + 1
    End If
    If IsDate(filtroFechaHasta) Then
        sql = sql & " AND S.FECHA_CREACION < ?"
        ReDim Preserve listaParams(n) : listaParams(n) = DateAdd("d", 1, CDate(filtroFechaHasta)) : n = n + 1
    End If

    sql = sql & " ORDER BY S.FECHA_CREACION DESC LIMIT " & CLng(offset) & ", " & CLng(limite)
    Set BuscarSolicitudes = EjecutarConsulta(conn, sql, listaParams)
End Function

Function ContarSolicitudes(conn, filtroCodSolicitud, filtroImei, filtroMsisdn, filtroTipo, filtroEstado, filtroIdentCliente, filtroNombreCliente, filtroFechaDesde, filtroFechaHasta)
    Dim sql, listaParams(), n, rs
    sql = "SELECT COUNT(*) AS N " & _
      "FROM SOLICITUD S " & _
      "INNER JOIN TIPO_SOLICITUD_REF T ON S.COD_TIPO_SOLICITUD = T.COD_TIPO_SOLICITUD " & _
      "INNER JOIN CLIENTE C ON S.COD_CLIENTE = C.COD_CLIENTE " & _
      "INNER JOIN HISTORICO_ESTADO_SOLICITUD H ON H.COD_INDICE_SOL = S.COD_INDICE_SOL " & _
      "AND H.FECHA_INICIO = ( " & _
      "SELECT MAX(H2.FECHA_INICIO) " & _
      "FROM HISTORICO_ESTADO_SOLICITUD H2 " & _
      "WHERE H2.COD_INDICE_SOL = S.COD_INDICE_SOL " & _
      ") " & _
      "INNER JOIN ESTADO_SOLICITUD_REF E ON H.COD_ESTADO = E.COD_ESTADO " & _
      "WHERE 1=1"
    ReDim listaParams(-1)
    n = 0

    If EsCadenaNoVacia(filtroCodSolicitud) Then
        sql = sql & " AND S.COD_SOLICITUD LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroCodSolicitud & "%" : n = n + 1
    End If
    If EsCadenaNoVacia(filtroImei) Then
        sql = sql & " AND S.IMEI LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroImei & "%" : n = n + 1
    End If
    If EsCadenaNoVacia(filtroMsisdn) Then
        sql = sql & " AND S.MSISDN LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroMsisdn & "%" : n = n + 1
    End If
    If EsCadenaNoVacia(filtroTipo) And filtroTipo <> "-1" And IsNumeric(filtroTipo) Then
        sql = sql & " AND S.COD_TIPO_SOLICITUD = ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CLng(filtroTipo) : n = n + 1
    End If
    If EsCadenaNoVacia(filtroEstado) And filtroEstado <> "-1" And IsNumeric(filtroEstado) Then
        sql = sql & " AND H.COD_ESTADO = ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CLng(filtroEstado) : n = n + 1
    End If
    If EsCadenaNoVacia(filtroIdentCliente) Then
        sql = sql & " AND C.NUM_IDENTIFICACION LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroIdentCliente & "%" : n = n + 1
    End If
    If EsCadenaNoVacia(filtroNombreCliente) Then
        sql = sql & " AND C.NOMBRE LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroNombreCliente & "%" : n = n + 1
    End If
    If IsDate(filtroFechaDesde) Then
        sql = sql & " AND S.FECHA_CREACION >= ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CDate(filtroFechaDesde) : n = n + 1
    End If
    If IsDate(filtroFechaHasta) Then
        sql = sql & " AND S.FECHA_CREACION < ?"
        ReDim Preserve listaParams(n) : listaParams(n) = DateAdd("d", 1, CDate(filtroFechaHasta)) : n = n + 1
    End If

    Set rs = EjecutarConsulta(conn, sql, listaParams)
    ContarSolicitudes = CLng(rs("N"))
    rs.Close
End Function

Function ObtenerSolicitudPorIndice(conn, codIndiceSol)
    Dim rs, s
    Set rs = EjecutarConsulta(conn, _
        "SELECT S.COD_INDICE_SOL, S.COD_SOLICITUD, S.COD_TIPO_SOLICITUD, T.DESCRIPCION AS TIPO, " & _
        "S.MSISDN, S.IMEI, S.NUM_DENUNCIA, S.FECHA_DENUNCIA, S.FECHA_ROBO, S.FECHA_SOLICITUD, S.FECHA_CREACION, " & _
        "CL.NOMBRE AS NOMBRE_CLIENTE, CL.PRIMER_APELLIDO, CL.SEGUNDO_APELLIDO, CL.NUM_IDENTIFICACION, " & _
        "U.COD_USUARIO AS USUARIO_CREADOR " & _
        "FROM SOLICITUD S " & _
        "INNER JOIN TIPO_SOLICITUD_REF T ON S.COD_TIPO_SOLICITUD = T.COD_TIPO_SOLICITUD " & _
        "INNER JOIN CLIENTE CL ON S.COD_CLIENTE = CL.COD_CLIENTE " & _
        "INNER JOIN USUARIOS U ON S.COD_INDICE_USUARIO = U.COD_INDICE_USUARIO " & _
        "WHERE S.COD_INDICE_SOL = ?", Array(codIndiceSol))
    If rs.EOF Then
        Set ObtenerSolicitudPorIndice = Nothing
    Else
        Set s = Server.CreateObject("Scripting.Dictionary")
        s("CodIndiceSol") = CLng(rs("COD_INDICE_SOL"))
        s("CodSolicitud") = rs("COD_SOLICITUD") & ""
        s("CodTipoSolicitud") = CLng(rs("COD_TIPO_SOLICITUD"))
        s("Tipo") = rs("TIPO") & ""
        s("Msisdn") = rs("MSISDN") & ""
        s("Imei") = rs("IMEI") & ""
        s("NumDenuncia") = rs("NUM_DENUNCIA") & ""
        s("FechaDenuncia") = rs("FECHA_DENUNCIA")
        s("FechaRobo") = rs("FECHA_ROBO")
        s("FechaSolicitud") = rs("FECHA_SOLICITUD")
        s("FechaCreacion") = rs("FECHA_CREACION")
        s("NombreCliente") = rs("NOMBRE_CLIENTE") & ""
        s("PrimerApellido") = rs("PRIMER_APELLIDO") & ""
        s("SegundoApellido") = rs("SEGUNDO_APELLIDO") & ""
        s("NumIdentificacion") = rs("NUM_IDENTIFICACION") & ""
        s("UsuarioCreador") = rs("USUARIO_CREADOR") & ""
        Set ObtenerSolicitudPorIndice = s
    End If
    rs.Close
End Function

Function ObtenerHistoricoEstados(conn, codIndiceSol)
    Set ObtenerHistoricoEstados = EjecutarConsulta(conn, _
        "SELECT E.DESCRIPCION, H.FECHA_INICIO, H.FECHA_FIN FROM HISTORICO_ESTADO_SOLICITUD H " & _
        "INNER JOIN ESTADO_SOLICITUD_REF E ON H.COD_ESTADO = E.COD_ESTADO " & _
        "WHERE H.COD_INDICE_SOL = ? ORDER BY H.FECHA_INICIO", Array(codIndiceSol))
End Function

Function ObtenerConfirmaciones(conn, codIndiceSol)
    Set ObtenerConfirmaciones = EjecutarConsulta(conn, _
        "SELECT O.NOMBRE AS OPERADOR, TC.DESCRIPCION AS TIPO_CONFIRMACION, CF.FECHA_CONFIRMACION " & _
        "FROM CONFIRMACION CF " & _
        "INNER JOIN OPERADOR_REF O ON CF.COD_OPERADOR = O.COD_OPERADOR " & _
        "INNER JOIN TIPO_CONFIRMACION_REF TC ON CF.COD_TIPO_CONFIRMACION = TC.COD_TIPO_CONFIRMACION " & _
        "WHERE CF.COD_INDICE_SOL = ? ORDER BY CF.FECHA_CONFIRMACION", Array(codIndiceSol))
End Function

Function ContarIncidenciasAsociadas(conn, codIndiceSol)
    Dim rs
    Set rs = EjecutarConsulta(conn, "SELECT COUNT(1) AS total FROM INCIDENCIA_TO_OBJETO WHERE COD_OBJETO = ? AND TIPO_OBJETO = 'S'", Array(codIndiceSol))
    ContarIncidenciasAsociadas = CLng(rs("total"))
    rs.Close
End Function

Function EstadoVigentePorIndice(conn, codIndiceSol)
    Dim rs
    Set rs = EjecutarConsulta(conn, "SELECT COD_ESTADO FROM HISTORICO_ESTADO_SOLICITUD WHERE COD_INDICE_SOL = ? AND FECHA_FIN IS NULL", Array(codIndiceSol))
    If rs.EOF Then
        EstadoVigentePorIndice = -1
    Else
        EstadoVigentePorIndice = CLng(rs("COD_ESTADO"))
    End If
    rs.Close
End Function

' Asuncion (no confirmada con precision en el codigo legacy disponible): una solicitud se
' considera cancelable mientras su estado vigente no sea uno de los 3 terminales, igual
' que el criterio de "en curso" usado para bloquear altas duplicadas sobre el mismo IMEI.
Function EsSolicitudCancelable(conn, codIndiceSol)
    Dim estadoVigente, codRechazada, codCancelada, codActualizada
    estadoVigente = EstadoVigentePorIndice(conn, codIndiceSol)
    codRechazada = CodigoEstadoSolicitud(conn, "Rechazada")
    codCancelada = CodigoEstadoSolicitud(conn, "Cancelada")
    codActualizada = CodigoEstadoSolicitud(conn, "Actualizada")
    EsSolicitudCancelable = (estadoVigente <> -1) And (estadoVigente <> codRechazada) And (estadoVigente <> codCancelada) And (estadoVigente <> codActualizada)
End Function

' Mismo patron confirmado en el legacy (cambioEstadoSolicitudes): cierra el estado
' vigente y abre uno nuevo, en lote y transaccional.
Function CancelarSolicitudesLote(conn, listaIds, codEstadoCancelada)
    Dim i, ok
    ok = True
    conn.BeginTrans
    On Error Resume Next
    For i = 0 To UBound(listaIds)
        Call EjecutarNonQueryTx(conn, "UPDATE HISTORICO_ESTADO_SOLICITUD SET FECHA_FIN = NOW() WHERE COD_INDICE_SOL = ? AND FECHA_FIN IS NULL", Array(listaIds(i)))
        If Err.Number <> 0 Then
            ok = False
            Exit For
        End If
        Call EjecutarNonQueryTx(conn, "INSERT INTO HISTORICO_ESTADO_SOLICITUD (COD_INDICE_SOL, COD_ESTADO, FECHA_INICIO) VALUES (?, ?, NOW())", Array(listaIds(i), codEstadoCancelada))
        If Err.Number <> 0 Then
            ok = False
            Exit For
        End If
    Next
    If Not ok Then
        conn.RollbackTrans
    Else
        conn.CommitTrans
    End If
    On Error Goto 0
    CancelarSolicitudesLote = ok
End Function
%>
