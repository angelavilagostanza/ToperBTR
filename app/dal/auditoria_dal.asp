<%
' DAL de consulta de AUDITORIA_ACCION (solo lectura - la escritura vive en
' app/include/security/auditoria.asp, RegistrarAuditoria, usada desde todos los modulos).
' COD_INDICE_USUARIO es nullable en la tabla (acciones sin usuario identificado: login
' fallido con usuario inexistente, el endpoint de intercambio con operadores externos sin
' sesion) - por eso el join a USUARIOS es siempre LEFT JOIN, nunca INNER.

' Catalogo dinamico de valores de ACCION ya usados (no hay tabla de referencia para esto,
' son literales de texto definidos en cada punto del codigo que llama a RegistrarAuditoria).
Function ListarAccionesAuditoria(conn)
    Set ListarAccionesAuditoria = EjecutarConsultaTx(conn, _
        "SELECT DISTINCT ACCION FROM AUDITORIA_ACCION ORDER BY ACCION", Array())
End Function

Function ContarAuditoria(conn, filtroUsuario, filtroAccion, filtroDesde, filtroHasta, filtroDetalle)
    Dim sql, listaParams(), n, rs
    sql = "SELECT COUNT(*) AS N FROM AUDITORIA_ACCION A LEFT JOIN USUARIOS U ON U.COD_INDICE_USUARIO = A.COD_INDICE_USUARIO WHERE 1=1"
    ReDim listaParams(-1)
    n = 0

    If EsCadenaNoVacia(filtroUsuario) Then
        sql = sql & " AND U.COD_USUARIO LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroUsuario & "%" : n = n + 1
    End If
    If EsCadenaNoVacia(filtroAccion) Then
        sql = sql & " AND A.ACCION = ?"
        ReDim Preserve listaParams(n) : listaParams(n) = filtroAccion : n = n + 1
    End If
    If IsDate(filtroDesde) Then
        sql = sql & " AND A.FECHA >= ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CDate(filtroDesde) : n = n + 1
    End If
    If IsDate(filtroHasta) Then
        sql = sql & " AND A.FECHA < DATE_ADD(?, INTERVAL 1 DAY)"
        ReDim Preserve listaParams(n) : listaParams(n) = CDate(filtroHasta) : n = n + 1
    End If
    ' El filtro de texto libre sobre DETALLE (VARCHAR(500), sin indice) solo se aplica si
    ' ya hay algun otro criterio que acote la busqueda - decision explicita para no permitir
    ' un LIKE '%...%' sin acotar sobre toda la tabla de auditoria.
    If EsCadenaNoVacia(filtroDetalle) And n > 0 Then
        sql = sql & " AND A.DETALLE LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroDetalle & "%" : n = n + 1
    End If

    Set rs = EjecutarConsulta(conn, sql, listaParams)
    ContarAuditoria = CLng(rs("N"))
    rs.Close
End Function

Function BuscarAuditoria(conn, filtroUsuario, filtroAccion, filtroDesde, filtroHasta, filtroDetalle, limite, offset)
    Dim sql, listaParams(), n
    sql = "SELECT A.FECHA, A.ACCION, A.ENTIDAD, A.ENTIDAD_ID, A.RESULTADO, A.DETALLE, A.IP, U.COD_USUARIO " & _
          "FROM AUDITORIA_ACCION A LEFT JOIN USUARIOS U ON U.COD_INDICE_USUARIO = A.COD_INDICE_USUARIO WHERE 1=1"
    ReDim listaParams(-1)
    n = 0

    If EsCadenaNoVacia(filtroUsuario) Then
        sql = sql & " AND U.COD_USUARIO LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroUsuario & "%" : n = n + 1
    End If
    If EsCadenaNoVacia(filtroAccion) Then
        sql = sql & " AND A.ACCION = ?"
        ReDim Preserve listaParams(n) : listaParams(n) = filtroAccion : n = n + 1
    End If
    If IsDate(filtroDesde) Then
        sql = sql & " AND A.FECHA >= ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CDate(filtroDesde) : n = n + 1
    End If
    If IsDate(filtroHasta) Then
        sql = sql & " AND A.FECHA < DATE_ADD(?, INTERVAL 1 DAY)"
        ReDim Preserve listaParams(n) : listaParams(n) = CDate(filtroHasta) : n = n + 1
    End If
    If EsCadenaNoVacia(filtroDetalle) And n > 0 Then
        sql = sql & " AND A.DETALLE LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroDetalle & "%" : n = n + 1
    End If

    sql = sql & " ORDER BY A.FECHA DESC LIMIT ? OFFSET ?"
    ReDim Preserve listaParams(n) : listaParams(n) = CLng(limite) : n = n + 1
    ReDim Preserve listaParams(n) : listaParams(n) = CLng(offset) : n = n + 1

    Set BuscarAuditoria = EjecutarConsulta(conn, sql, listaParams)
End Function
%>
