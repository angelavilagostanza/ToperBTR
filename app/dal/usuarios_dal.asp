<%
Function BuscarUsuarios(conn, filtroLogin, filtroNombre, filtroApellidos, filtroPerfil, filtroEstado)
    Dim sql, listaParams(), n
    sql = "SELECT U.COD_INDICE_USUARIO, U.COD_USUARIO, U.NOMBRE, U.APELLIDOS, U.COD_PERFIL, P.NOMBRE AS NOMBRE_PERFIL, U.ESTADO, E.DESCRIPCION AS DESCRIPCION_ESTADO " & _
          "FROM USUARIOS U " & _
          "INNER JOIN PERFIL_REF P ON U.COD_PERFIL = P.COD_PERFIL " & _
          "INNER JOIN ESTADO_USUARIO_REF E ON U.ESTADO = E.ESTADO " & _
          "WHERE 1=1"
    ReDim listaParams(-1)
    n = 0

    If EsCadenaNoVacia(filtroLogin) Then
        sql = sql & " AND U.COD_USUARIO LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroLogin & "%" : n = n + 1
    End If
    If EsCadenaNoVacia(filtroNombre) Then
        sql = sql & " AND U.NOMBRE LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroNombre & "%" : n = n + 1
    End If
    If EsCadenaNoVacia(filtroApellidos) Then
        sql = sql & " AND U.APELLIDOS LIKE ?"
        ReDim Preserve listaParams(n) : listaParams(n) = "%" & filtroApellidos & "%" : n = n + 1
    End If
    If EsCadenaNoVacia(filtroPerfil) And filtroPerfil <> "-1" And IsNumeric(filtroPerfil) Then
        sql = sql & " AND U.COD_PERFIL = ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CLng(filtroPerfil) : n = n + 1
    End If
    If EsCadenaNoVacia(filtroEstado) And filtroEstado <> "-1" And IsNumeric(filtroEstado) Then
        sql = sql & " AND U.ESTADO = ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CLng(filtroEstado) : n = n + 1
    End If

    sql = sql & " ORDER BY U.COD_USUARIO"

    Set BuscarUsuarios = EjecutarConsulta(conn, sql, listaParams)
End Function

Function ObtenerUsuarioPorIndice(conn, codIndiceUsuario)
    Dim rs, u
    Set rs = EjecutarConsulta(conn, _
        "SELECT COD_INDICE_USUARIO, COD_USUARIO, NOMBRE, APELLIDOS, PASSWORD, ALGORITMO_PASSWORD, FECHA_CADUCIDAD, NUM_FALLOS, ESTADO, COD_PERFIL, COD_OPERADOR FROM USUARIOS WHERE COD_INDICE_USUARIO = ?", _
        Array(codIndiceUsuario))
    If rs.EOF Then
        Set ObtenerUsuarioPorIndice = Nothing
    Else
        Set u = Server.CreateObject("Scripting.Dictionary")
        u("CodIndiceUsuario") = CLng(rs("COD_INDICE_USUARIO"))
        u("CodUsuario") = rs("COD_USUARIO") & ""
        u("Nombre") = rs("NOMBRE") & ""
        u("Apellidos") = rs("APELLIDOS") & ""
        u("Password") = rs("PASSWORD") & ""
        u("AlgoritmoPassword") = rs("ALGORITMO_PASSWORD") & ""
        u("FechaCaducidad") = rs("FECHA_CADUCIDAD")
        u("NumFallos") = CLng(rs("NUM_FALLOS"))
        u("Estado") = CLng(rs("ESTADO"))
        u("CodPerfil") = CLng(rs("COD_PERFIL"))
        u("CodOperador") = CLng(rs("COD_OPERADOR"))
        Set ObtenerUsuarioPorIndice = u
    End If
    rs.Close
End Function

Function ExisteLogin(conn, codUsuario)
    Dim rs
    Set rs = EjecutarConsulta(conn, "SELECT COD_INDICE_USUARIO FROM USUARIOS WHERE COD_USUARIO = ?", Array(codUsuario))
    ExisteLogin = Not rs.EOF
    rs.Close
End Function

Function ContarUsuariosActivosOperador(conn, codOperador, codEstadoBorrado)
    Dim rs
    Set rs = EjecutarConsulta(conn, "SELECT COUNT(*) AS total FROM USUARIOS WHERE COD_OPERADOR = ? AND ESTADO <> ?", Array(codOperador, codEstadoBorrado))
    ContarUsuariosActivosOperador = CLng(rs("total"))
    rs.Close
End Function

Function NumMaxUsuariosOperador(conn, codOperador)
    Dim rs
    Set rs = EjecutarConsulta(conn, "SELECT NUM_MAX_USUARIOS FROM OPERADOR_REF WHERE COD_OPERADOR = ?", Array(codOperador))
    If rs.EOF Then
        NumMaxUsuariosOperador = 0
    Else
        NumMaxUsuariosOperador = CLng(rs("NUM_MAX_USUARIOS"))
    End If
    rs.Close
End Function

Sub CrearUsuario(conn, codUsuario, nombre, apellidos, hashPassword, algoritmo, fechaCaducidad, codPerfil, codOperador, codEstadoActivo)
    Call EjecutarNonQuery(conn, _
        "INSERT INTO USUARIOS (COD_USUARIO, NOMBRE, APELLIDOS, PASSWORD, ALGORITMO_PASSWORD, FECHA_CADUCIDAD, FECHA_ALTA, NUM_FALLOS, FECHA_ULTIMA_CONEXION, ESTADO, COD_PERFIL, COD_OPERADOR) " & _
        "VALUES (?, ?, ?, ?, ?, ?, NOW(), 0, NULL, ?, ?, ?)", _
        Array(codUsuario, nombre, apellidos, hashPassword, algoritmo, fechaCaducidad, codEstadoActivo, codPerfil, codOperador))
End Sub

' Baja y desbloqueo se aplican SIEMPRE en lote (aunque sea un unico elemento), reproduciendo
' el UNICO patron del legacy que funciona de verdad (cambioEstadoUsuarios por indice,
' transaccional) y nunca el patron "por cod_usuario individual" que en el legacy quedo
' inutilizado desde 2011 (ver docs/toperbtr-java-findings en memoria del proyecto).
Function DarBajaUsuariosLote(conn, listaIds, codEstadoBorrado)
    Dim i, ok
    ok = True
    conn.BeginTrans
    On Error Resume Next
    For i = 0 To UBound(listaIds)
        Call EjecutarNonQueryTx(conn, "UPDATE USUARIOS SET ESTADO = ? WHERE COD_INDICE_USUARIO = ? AND ESTADO <> ?", Array(codEstadoBorrado, listaIds(i), codEstadoBorrado))
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
    DarBajaUsuariosLote = ok
End Function

Function DesbloquearUsuariosLote(conn, listaIds, codEstadoActivo)
    Dim i, ok
    ok = True
    conn.BeginTrans
    On Error Resume Next
    For i = 0 To UBound(listaIds)
        Call EjecutarNonQueryTx(conn, "UPDATE USUARIOS SET ESTADO = ?, NUM_FALLOS = 0 WHERE COD_INDICE_USUARIO = ?", Array(codEstadoActivo, listaIds(i)))
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
    DesbloquearUsuariosLote = ok
End Function

' Se llama siempre desde dentro de la transaccion de CambiarPasswordConHistorico (via
' RotarHistoricoPassword) - usa la variante Tx por el mismo motivo que las demas de aqui.
Function ContarHistoricoPassword(conn, codIndiceUsuario)
    Dim rs
    Set rs = EjecutarConsultaTx(conn, "SELECT COUNT(*) AS total FROM HISTORICO_PASSWORD WHERE COD_INDICE_USUARIO = ?", Array(codIndiceUsuario))
    ContarHistoricoPassword = CLng(rs("total"))
    rs.Close
End Function

' Comprueba la contrasena nueva contra la actual Y contra el historico (hasta NUM_OLD_PASSWORD
' filas). Cada fila del historico puede estar en formato legacy (SHA1B64) o ya migrado
' (PBKDF2SHA256) - VerificarPassword discrimina por su propia columna ALGORITMO_PASSWORD.
Function PasswordReutilizada(conn, codIndiceUsuario, passwordNuevaPlano, passwordActualHash, algoritmoActual)
    Dim reutilizada, rs
    reutilizada = False

    If VerificarPassword(passwordNuevaPlano, passwordActualHash, algoritmoActual) Then
        PasswordReutilizada = True
        Exit Function
    End If

    Set rs = EjecutarConsulta(conn, "SELECT PASSWORD, ALGORITMO_PASSWORD FROM HISTORICO_PASSWORD WHERE COD_INDICE_USUARIO = ?", Array(codIndiceUsuario))
    Do While Not rs.EOF
        If VerificarPassword(passwordNuevaPlano, rs("PASSWORD") & "", rs("ALGORITMO_PASSWORD") & "") Then
            reutilizada = True
            Exit Do
        End If
        rs.MoveNext
    Loop
    rs.Close
    PasswordReutilizada = reutilizada
End Function

' Mantiene siempre como maximo maxHistorico filas por usuario: mientras no se alcanza el
' limite, inserta; al llegar al limite, recicla la fila con FECHA_CADUCO mas antigua.
' Decision de negocio: maxHistorico = 5 (ver docs/toperbtr-decisions en memoria del proyecto).
Sub RotarHistoricoPassword(conn, codIndiceUsuario, passwordSaliente, algoritmoSaliente, maxHistorico)
    Dim total
    total = ContarHistoricoPassword(conn, codIndiceUsuario)
    If total < maxHistorico Then
        Call EjecutarNonQueryTx(conn, _
            "INSERT INTO HISTORICO_PASSWORD (COD_INDICE_USUARIO, PASSWORD, ALGORITMO_PASSWORD, FECHA_CADUCO) VALUES (?, ?, ?, NOW())", _
            Array(codIndiceUsuario, passwordSaliente, algoritmoSaliente))
    Else
        Call EjecutarNonQueryTx(conn, _
            "UPDATE HISTORICO_PASSWORD SET PASSWORD = ?, ALGORITMO_PASSWORD = ?, FECHA_CADUCO = NOW() " & _
            "WHERE COD_INDICE_USUARIO = ? AND FECHA_CADUCO = (SELECT MINIMO FROM (SELECT MIN(FECHA_CADUCO) AS MINIMO FROM HISTORICO_PASSWORD WHERE COD_INDICE_USUARIO = ?) AS AUX)", _
            Array(passwordSaliente, algoritmoSaliente, codIndiceUsuario, codIndiceUsuario))
    End If
End Sub

' Se llama siempre desde dentro de la transaccion de CambiarPasswordConHistorico - usa la
' variante Tx (ver nota al principio del fichero de conexion sobre por que).
Sub ActualizarPassword(conn, codIndiceUsuario, nuevoHash, fechaCaducidad)
    Call EjecutarNonQueryTx(conn, _
        "UPDATE USUARIOS SET PASSWORD = ?, ALGORITMO_PASSWORD = 'PBKDF2SHA256', FECHA_CADUCIDAD = ?, NUM_FALLOS = 0 WHERE COD_INDICE_USUARIO = ?", _
        Array(nuevoHash, fechaCaducidad, codIndiceUsuario))
End Sub

' Punto de entrada unico para cambiar una contrasena, usado tanto por el autoservicio
' (cambio_password_do.asp) como por el reset de perfil Seguridad (reset_password_admin_do.asp).
' A diferencia del legacy, AMBOS caminos comprueban reutilizacion y rotan el historico igual
' (decision de negocio: "cambioPasswordAdm no puede reutilizar contrasenas").
' Devuelve "" si todo fue bien, o un mensaje de error para mostrar al usuario.
Function CambiarPasswordConHistorico(conn, codIndiceUsuario, passwordActualHash, algoritmoActual, passwordNuevaPlano, maxHistorico, diasCaducidad)
    If PasswordReutilizada(conn, codIndiceUsuario, passwordNuevaPlano, passwordActualHash, algoritmoActual) Then
        CambiarPasswordConHistorico = "La contrasena ya se ha utilizado anteriormente. Elija una distinta."
        Exit Function
    End If

    Dim nuevoHash, fechaCaducidad, huboError
    nuevoHash = CrearHashPbkdf2(passwordNuevaPlano, CLng(Application("PBKDF2Iteraciones")))
    fechaCaducidad = DateAdd("d", diasCaducidad, Now())
    huboError = False

    conn.BeginTrans
    On Error Resume Next
    Call RotarHistoricoPassword(conn, codIndiceUsuario, passwordActualHash, algoritmoActual, maxHistorico)
    If Err.Number <> 0 Then huboError = True
    If Not huboError Then
        Call ActualizarPassword(conn, codIndiceUsuario, nuevoHash, fechaCaducidad)
        If Err.Number <> 0 Then huboError = True
    End If

    If huboError Then
        conn.RollbackTrans
        CambiarPasswordConHistorico = MensajeErrorGenerico()
    Else
        conn.CommitTrans
        CambiarPasswordConHistorico = ""
    End If
    On Error Goto 0
End Function
%>
