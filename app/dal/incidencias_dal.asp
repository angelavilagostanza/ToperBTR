<%
' El listado de incidencias del legacy tenia un bug real confirmado en el codigo Java:
' los joins a SOLICITUD/FICHEROS eran condicionales segun el filtro activo, pero NUNCA
' discriminaban por INCIDENCIA_TO_OBJETO.TIPO_OBJETO - podia ocultar o mezclar incidencias
' si los rangos de ID de solicitud y de fichero coincidian por casualidad. Aqui se corrige
' con LEFT JOIN discriminando el TIPO_OBJETO dentro de la propia condicion de join (no en
' el WHERE), asi cada incidencia solo casa con la tabla que le corresponde de verdad.
Function BuscarIncidencias(conn, filtroCodIncidencia, filtroTipo, filtroEstado, filtroOperador, filtroCodSolicitud, filtroNombreFichero, filtroFechaDesde, filtroFechaHasta, limite, offset)
    Dim sql, listaParams(), n
    sql = "SELECT CAST(I.COD_IND_INCIDENCIA AS CHAR) AS COD_IND_INCIDENCIA, I.COD_INCIDENCIA, I.FECHA_CREACION, I.FECHA_RESOLUCION, " & _
          "TI.DESCRIPCION AS TIPO_INCIDENCIA, OP.NOMBRE AS OPERADOR, E.DESCRIPCION AS ESTADO, " & _
          "O.TIPO_OBJETO, O.COD_OBJETO, S.COD_SOLICITUD, F.NOMBRE AS NOMBRE_FICHERO " & _
          "FROM INCIDENCIA I " & _
          "INNER JOIN TIPO_INCIDENCIA_REF TI ON I.COD_TIPO_INCIDENCIA = TI.COD_TIPO_INCIDENCIA " & _
          "INNER JOIN OPERADOR_REF OP ON I.COD_OPERADOR = OP.COD_OPERADOR " & _
          "INNER JOIN HISTORICO_ESTADO_INCIDENCIA H ON H.COD_IND_INCIDENCIA = I.COD_IND_INCIDENCIA AND H.FECHA_FIN IS NULL " & _
          "INNER JOIN ESTADO_INCIDENCIA_REF E ON H.COD_ESTADO = E.COD_ESTADO " & _
          "INNER JOIN INCIDENCIA_TO_OBJETO O ON O.COD_IND_INCIDENCIA = I.COD_IND_INCIDENCIA " & _
          "LEFT JOIN SOLICITUD S ON O.TIPO_OBJETO = 'S' AND O.COD_OBJETO = S.COD_INDICE_SOL " & _
          "LEFT JOIN FICHEROS F ON O.TIPO_OBJETO = 'F' AND O.COD_OBJETO = F.COD_FICHERO " & _
          "WHERE 1=1"
    ReDim listaParams(-1)
    n = 0

    If EsCadenaNoVacia(filtroCodIncidencia) Then
        sql = sql & " AND I.COD_INCIDENCIA LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroCodIncidencia & "%" : n = n + 1
    End If
    If EsCadenaNoVacia(filtroTipo) And filtroTipo <> "-1" And IsNumeric(filtroTipo) Then
        sql = sql & " AND I.COD_TIPO_INCIDENCIA = ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CLng(filtroTipo) : n = n + 1
    End If
    If EsCadenaNoVacia(filtroEstado) And filtroEstado <> "-1" And IsNumeric(filtroEstado) Then
        sql = sql & " AND H.COD_ESTADO = ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CLng(filtroEstado) : n = n + 1
    End If
    If EsCadenaNoVacia(filtroOperador) And filtroOperador <> "-1" And IsNumeric(filtroOperador) Then
        sql = sql & " AND I.COD_OPERADOR = ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CLng(filtroOperador) : n = n + 1
    End If
    If EsCadenaNoVacia(filtroCodSolicitud) Then
        sql = sql & " AND S.COD_SOLICITUD LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroCodSolicitud & "%" : n = n + 1
    End If
    If EsCadenaNoVacia(filtroNombreFichero) Then
        sql = sql & " AND F.NOMBRE LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroNombreFichero & "%" : n = n + 1
    End If
    If IsDate(filtroFechaDesde) Then
        sql = sql & " AND I.FECHA_CREACION >= ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CDate(filtroFechaDesde) : n = n + 1
    End If
    If IsDate(filtroFechaHasta) Then
        sql = sql & " AND I.FECHA_CREACION < ?"
        ReDim Preserve listaParams(n) : listaParams(n) = DateAdd("d", 1, CDate(filtroFechaHasta)) : n = n + 1
    End If

    sql = sql & " ORDER BY I.FECHA_CREACION DESC LIMIT " & CLng(offset) & ", " & CLng(limite)
    Set BuscarIncidencias = EjecutarConsulta(conn, sql, listaParams)
End Function

Function ContarIncidencias(conn, filtroCodIncidencia, filtroTipo, filtroEstado, filtroOperador, filtroCodSolicitud, filtroNombreFichero, filtroFechaDesde, filtroFechaHasta)
    Dim sql, listaParams(), n, rs
    sql = "SELECT COUNT(*) AS N " & _
          "FROM INCIDENCIA I " & _
          "INNER JOIN TIPO_INCIDENCIA_REF TI ON I.COD_TIPO_INCIDENCIA = TI.COD_TIPO_INCIDENCIA " & _
          "INNER JOIN OPERADOR_REF OP ON I.COD_OPERADOR = OP.COD_OPERADOR " & _
          "INNER JOIN HISTORICO_ESTADO_INCIDENCIA H ON H.COD_IND_INCIDENCIA = I.COD_IND_INCIDENCIA AND H.FECHA_FIN IS NULL " & _
          "INNER JOIN ESTADO_INCIDENCIA_REF E ON H.COD_ESTADO = E.COD_ESTADO " & _
          "INNER JOIN INCIDENCIA_TO_OBJETO O ON O.COD_IND_INCIDENCIA = I.COD_IND_INCIDENCIA " & _
          "LEFT JOIN SOLICITUD S ON O.TIPO_OBJETO = 'S' AND O.COD_OBJETO = S.COD_INDICE_SOL " & _
          "LEFT JOIN FICHEROS F ON O.TIPO_OBJETO = 'F' AND O.COD_OBJETO = F.COD_FICHERO " & _
          "WHERE 1=1"
    ReDim listaParams(-1)
    n = 0

    If EsCadenaNoVacia(filtroCodIncidencia) Then
        sql = sql & " AND I.COD_INCIDENCIA LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroCodIncidencia & "%" : n = n + 1
    End If
    If EsCadenaNoVacia(filtroTipo) And filtroTipo <> "-1" And IsNumeric(filtroTipo) Then
        sql = sql & " AND I.COD_TIPO_INCIDENCIA = ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CLng(filtroTipo) : n = n + 1
    End If
    If EsCadenaNoVacia(filtroEstado) And filtroEstado <> "-1" And IsNumeric(filtroEstado) Then
        sql = sql & " AND H.COD_ESTADO = ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CLng(filtroEstado) : n = n + 1
    End If
    If EsCadenaNoVacia(filtroOperador) And filtroOperador <> "-1" And IsNumeric(filtroOperador) Then
        sql = sql & " AND I.COD_OPERADOR = ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CLng(filtroOperador) : n = n + 1
    End If
    If EsCadenaNoVacia(filtroCodSolicitud) Then
        sql = sql & " AND S.COD_SOLICITUD LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroCodSolicitud & "%" : n = n + 1
    End If
    If EsCadenaNoVacia(filtroNombreFichero) Then
        sql = sql & " AND F.NOMBRE LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroNombreFichero & "%" : n = n + 1
    End If
    If IsDate(filtroFechaDesde) Then
        sql = sql & " AND I.FECHA_CREACION >= ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CDate(filtroFechaDesde) : n = n + 1
    End If
    If IsDate(filtroFechaHasta) Then
        sql = sql & " AND I.FECHA_CREACION < ?"
        ReDim Preserve listaParams(n) : listaParams(n) = DateAdd("d", 1, CDate(filtroFechaHasta)) : n = n + 1
    End If

    Set rs = EjecutarConsulta(conn, sql, listaParams)
    ContarIncidencias = CLng(rs("N"))
    rs.Close
End Function

Function ObtenerIncidenciaPorIndice(conn, codIndIncidencia)
    Dim rs, inc
    Set rs = EjecutarConsulta(conn, _
        "SELECT CAST(I.COD_IND_INCIDENCIA AS CHAR) AS COD_IND_INCIDENCIA, I.COD_INCIDENCIA, I.OTROS_DATOS, I.CAUSA, I.FECHA_CREACION, I.FECHA_RESOLUCION, " & _
        "TI.DESCRIPCION AS TIPO_INCIDENCIA, TI.CATEGORIA, OP.NOMBRE AS OPERADOR, " & _
        "O.TIPO_OBJETO, CAST(O.COD_OBJETO AS CHAR) AS COD_OBJETO, S.COD_SOLICITUD, F.NOMBRE AS NOMBRE_FICHERO " & _
        "FROM INCIDENCIA I " & _
        "INNER JOIN TIPO_INCIDENCIA_REF TI ON I.COD_TIPO_INCIDENCIA = TI.COD_TIPO_INCIDENCIA " & _
        "INNER JOIN OPERADOR_REF OP ON I.COD_OPERADOR = OP.COD_OPERADOR " & _
        "INNER JOIN INCIDENCIA_TO_OBJETO O ON O.COD_IND_INCIDENCIA = I.COD_IND_INCIDENCIA " & _
        "LEFT JOIN SOLICITUD S ON O.TIPO_OBJETO = 'S' AND O.COD_OBJETO = S.COD_INDICE_SOL " & _
        "LEFT JOIN FICHEROS F ON O.TIPO_OBJETO = 'F' AND O.COD_OBJETO = F.COD_FICHERO " & _
        "WHERE I.COD_IND_INCIDENCIA = ?", Array(codIndIncidencia))
    If rs.EOF Then
        Set ObtenerIncidenciaPorIndice = Nothing
    Else
        Set inc = Server.CreateObject("Scripting.Dictionary")
        inc("CodIndIncidencia") = CLng(rs("COD_IND_INCIDENCIA"))
        inc("CodIncidencia") = rs("COD_INCIDENCIA") & ""
        inc("OtrosDatos") = rs("OTROS_DATOS") & ""
        inc("Causa") = rs("CAUSA") & ""
        inc("FechaCreacion") = rs("FECHA_CREACION")
        inc("FechaResolucion") = rs("FECHA_RESOLUCION")
        inc("TipoIncidencia") = rs("TIPO_INCIDENCIA") & ""
        inc("Categoria") = rs("CATEGORIA") & ""
        inc("Operador") = rs("OPERADOR") & ""
        inc("TipoObjeto") = rs("TIPO_OBJETO") & ""
        inc("CodObjeto") = CLng(rs("COD_OBJETO"))
        inc("CodSolicitud") = rs("COD_SOLICITUD") & ""
        inc("NombreFichero") = rs("NOMBRE_FICHERO") & ""
        Set ObtenerIncidenciaPorIndice = inc
    End If
    rs.Close
End Function

Function EstadoVigenteIncidencia(conn, codIndIncidencia)
    Dim rs
    Set rs = EjecutarConsulta(conn, _
        "SELECT E.DESCRIPCION FROM HISTORICO_ESTADO_INCIDENCIA H INNER JOIN ESTADO_INCIDENCIA_REF E ON H.COD_ESTADO = E.COD_ESTADO " & _
        "WHERE H.COD_IND_INCIDENCIA = ? AND H.FECHA_FIN IS NULL", Array(codIndIncidencia))
    If rs.EOF Then
        EstadoVigenteIncidencia = ""
    Else
        EstadoVigenteIncidencia = rs("DESCRIPCION") & ""
    End If
    rs.Close
End Function

Function ObtenerComentarios(conn, codIndIncidencia)
    Set ObtenerComentarios = EjecutarConsulta(conn, _
        "SELECT C.COMENTARIO, C.FECHA_CREACION, U.COD_USUARIO FROM COMENTARIO_INCIDENCIA C " & _
        "INNER JOIN USUARIOS U ON C.COD_INDICE_USUARIO = U.COD_INDICE_USUARIO " & _
        "WHERE C.COD_IND_INCIDENCIA = ? ORDER BY C.FECHA_CREACION", Array(codIndIncidencia))
End Function

Sub InsertarComentario(conn, codIndIncidencia, codIndiceUsuario, comentario)
    Call EjecutarNonQuery(conn, _
        "INSERT INTO COMENTARIO_INCIDENCIA (COD_INDICE_USUARIO, FECHA_CREACION, COMENTARIO, COD_IND_INCIDENCIA) VALUES (?, NOW(), ?, ?)", _
        Array(codIndiceUsuario, comentario, codIndIncidencia))
End Sub
%>
