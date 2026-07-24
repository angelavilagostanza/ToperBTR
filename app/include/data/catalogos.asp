<%
' Resolucion de catalogos siempre por NOMBRE/DESCRIPCION, nunca por literal numerico,
' y siempre consultando directamente (sin cache de aplicacion, decision de diseno de ToperBTR).
'
' Usan EjecutarConsultaTx (no la variante segura EjecutarConsulta) porque estas funciones
' se llaman indistintamente desde paginas de solo lectura Y desde dentro de transacciones
' ya abiertas por otras funciones (p.ej. CrearSolicitudCompleta) - si usaran la variante
' segura, un fallo aqui redirigiria y cortaria la respuesta antes de que la transaccion
' que las llamo pudiera hacer su propio RollbackTrans. Contrapartida aceptada: si una de
' estas consultas (simples, sobre tablas de catalogo pequenas) fallara en una pagina que
' NO esta dentro de ninguna transaccion, quedaria sin el aviso amigable automatico. Riesgo
' bajo dado lo simples que son estas consultas.

Function CodigoPerfil(conn, nombrePerfil)
    Dim rs
    Set rs = EjecutarConsultaTx(conn, "SELECT COD_PERFIL FROM PERFIL_REF WHERE NOMBRE = ?", Array(nombrePerfil))
    If rs.EOF Then
        CodigoPerfil = -1
    Else
        CodigoPerfil = CLng(rs("COD_PERFIL"))
    End If
    rs.Close
End Function

Function CodigoEstadoUsuario(conn, descripcion)
    Dim rs
    Set rs = EjecutarConsultaTx(conn, "SELECT ESTADO FROM ESTADO_USUARIO_REF WHERE DESCRIPCION = ?", Array(descripcion))
    If rs.EOF Then
        CodigoEstadoUsuario = -1
    Else
        CodigoEstadoUsuario = CLng(rs("ESTADO"))
    End If
    rs.Close
End Function

Function CodigoAccion(conn, nombre)
    Dim rs
    Set rs = EjecutarConsultaTx(conn, "SELECT COD_ACCION FROM ACCION_REF WHERE NOMBRE = ?", Array(nombre))
    If rs.EOF Then
        CodigoAccion = -1
    Else
        CodigoAccion = CLng(rs("COD_ACCION"))
    End If
    rs.Close
End Function

Function AccionesPermitidas(conn, codPerfil)
    Dim rs, resultado(), total
    ReDim resultado(-1)
    total = 0
    Set rs = EjecutarConsultaTx(conn, "SELECT COD_ACCION FROM ACCION_PERMITIDA_REF WHERE COD_PERFIL = ?", Array(codPerfil))
    Do While Not rs.EOF
        ReDim Preserve resultado(total)
        resultado(total) = CLng(rs("COD_ACCION"))
        total = total + 1
        rs.MoveNext
    Loop
    rs.Close
    AccionesPermitidas = resultado
End Function

Function TienePermiso(accionesPermitidas, codAccionRequerida)
    Dim i
    If codAccionRequerida = 0 Then
        TienePermiso = True
        Exit Function
    End If
    TienePermiso = False
    If Not IsArray(accionesPermitidas) Then Exit Function
    If UBound(accionesPermitidas) < 0 Then Exit Function
    For i = 0 To UBound(accionesPermitidas)
        If accionesPermitidas(i) = codAccionRequerida Then
            TienePermiso = True
            Exit Function
        End If
    Next
End Function

Function ValorParametro(conn, clave)
    Dim rs
    Set rs = EjecutarConsultaTx(conn, "SELECT VALOR FROM PARAMETROS_REF WHERE CLAVE = ?", Array(clave))
    If rs.EOF Then
        ValorParametro = ""
    Else
        ValorParametro = rs("VALOR") & ""
    End If
    rs.Close
End Function

Function ValorParametroEntero(conn, clave, valorPorDefecto)
    Dim v
    v = ValorParametro(conn, clave)
    If IsNumeric(v) Then
        ValorParametroEntero = CLng(v)
    Else
        ValorParametroEntero = valorPorDefecto
    End If
End Function

Function ListarPerfiles(conn)
    Set ListarPerfiles = EjecutarConsultaTx(conn, "SELECT COD_PERFIL, NOMBRE FROM PERFIL_REF ORDER BY NOMBRE", Array())
End Function

Function ListarEstadosUsuario(conn)
    Set ListarEstadosUsuario = EjecutarConsultaTx(conn, "SELECT ESTADO, DESCRIPCION FROM ESTADO_USUARIO_REF ORDER BY ESTADO", Array())
End Function

Function CodigoTipoSolicitud(conn, descripcion)
    Dim rs
    Set rs = EjecutarConsultaTx(conn, "SELECT COD_TIPO_SOLICITUD FROM TIPO_SOLICITUD_REF WHERE DESCRIPCION = ?", Array(descripcion))
    If rs.EOF Then
        CodigoTipoSolicitud = -1
    Else
        CodigoTipoSolicitud = CLng(rs("COD_TIPO_SOLICITUD"))
    End If
    rs.Close
End Function

' El legacy tenia una inconsistencia real: un punto del codigo resolvia "Inclusion" por
' descripcion y otro la hardcodeaba con un literal numerico distinto. Aqui SIEMPRE se
' resuelve por catalogo (por CLAVE para diferenciar Inclusion/Exclusion, o por
' DESCRIPCION segun convenga), nunca con un numero fijo.
Function CodigoTipoSolicitudPorClave(conn, clave)
    Dim rs
    Set rs = EjecutarConsultaTx(conn, "SELECT COD_TIPO_SOLICITUD FROM TIPO_SOLICITUD_REF WHERE CLAVE = ?", Array(clave))
    If rs.EOF Then
        CodigoTipoSolicitudPorClave = -1
    Else
        CodigoTipoSolicitudPorClave = CLng(rs("COD_TIPO_SOLICITUD"))
    End If
    rs.Close
End Function

Function ListarTiposSolicitud(conn)
    Set ListarTiposSolicitud = EjecutarConsultaTx(conn, "SELECT COD_TIPO_SOLICITUD, DESCRIPCION, CLAVE FROM TIPO_SOLICITUD_REF ORDER BY DESCRIPCION", Array())
End Function

Function CodigoEstadoSolicitud(conn, descripcion)
    Dim rs
    Set rs = EjecutarConsultaTx(conn, "SELECT COD_ESTADO FROM ESTADO_SOLICITUD_REF WHERE DESCRIPCION = ?", Array(descripcion))
    If rs.EOF Then
        CodigoEstadoSolicitud = -1
    Else
        CodigoEstadoSolicitud = CLng(rs("COD_ESTADO"))
    End If
    rs.Close
End Function

Function ListarEstadosSolicitud(conn)
    Set ListarEstadosSolicitud = EjecutarConsultaTx(conn, "SELECT COD_ESTADO, DESCRIPCION FROM ESTADO_SOLICITUD_REF ORDER BY COD_ESTADO", Array())
End Function

Function CodigoTipoBloqueo(conn, descripcion)
    Dim rs
    Set rs = EjecutarConsultaTx(conn, "SELECT COD_TIPO_BLOQUEO FROM TIPO_BLOQUEO_REF WHERE DESCRIPCION = ?", Array(descripcion))
    If rs.EOF Then
        CodigoTipoBloqueo = -1
    Else
        CodigoTipoBloqueo = CLng(rs("COD_TIPO_BLOQUEO"))
    End If
    rs.Close
End Function

Function ListarTiposIncidencia(conn)
    Set ListarTiposIncidencia = EjecutarConsultaTx(conn, "SELECT COD_TIPO_INCIDENCIA, DESCRIPCION FROM TIPO_INCIDENCIA_REF ORDER BY DESCRIPCION", Array())
End Function

Function ListarEstadosIncidencia(conn)
    Set ListarEstadosIncidencia = EjecutarConsultaTx(conn, "SELECT COD_ESTADO, DESCRIPCION FROM ESTADO_INCIDENCIA_REF ORDER BY COD_ESTADO", Array())
End Function

' Incluye CLAVE (usada por Ficheros para inferir el operador propietario a partir del
' nombre de fichero, igual que hacia el codigo legacy real) - anadir la columna no rompe
' a los consumidores existentes (Incidencias, Solicitudes) que solo leen COD_OPERADOR/NOMBRE.
Function ListarOperadores(conn)
    Set ListarOperadores = EjecutarConsultaTx(conn, "SELECT COD_OPERADOR, NOMBRE, CLAVE FROM OPERADOR_REF ORDER BY NOMBRE", Array())
End Function
%>
