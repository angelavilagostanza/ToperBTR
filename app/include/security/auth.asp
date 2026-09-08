<%
Sub RegistrarLoginCorrecto(conn, codIndiceUsuario)
    Call EjecutarNonQuery(conn, "UPDATE USUARIOS SET NUM_FALLOS = 0, FECHA_ULTIMA_CONEXION = NOW() WHERE COD_INDICE_USUARIO = ?", Array(codIndiceUsuario))
End Sub

Sub RehashPassword(conn, codIndiceUsuario, passwordPlano)
    Dim iteraciones, nuevoHash
    iteraciones = CLng(Application("PBKDF2Iteraciones"))
    nuevoHash = CrearHashPbkdf2(passwordPlano, iteraciones)
    Call EjecutarNonQuery(conn, "UPDATE USUARIOS SET PASSWORD = ?, ALGORITMO_PASSWORD = 'PBKDF2SHA256' WHERE COD_INDICE_USUARIO = ?", Array(nuevoHash, codIndiceUsuario))
End Sub

Function MensajeIntentosRestantes(conn, numFallosActual)
    Dim umbral, restantes
    umbral = ValorParametroEntero(conn, "NUM_MAX_INTENTOS_FALLIDOS", 5)
    restantes = umbral - numFallosActual
    If restantes <= 0 Then
        MensajeIntentosRestantes = "Su cuenta ha sido bloqueada por superar el numero de intentos fallidos. Contacte con el administrador de seguridad."
    ElseIf restantes = 1 Then
        MensajeIntentosRestantes = "Usuario o contrasena incorrectos. Le queda 1 intento antes de que su cuenta sea bloqueada."
    ElseIf restantes = 2 Then
        MensajeIntentosRestantes = "Usuario o contrasena incorrectos. Le quedan 2 intentos antes de que su cuenta sea bloqueada."
    Else
        MensajeIntentosRestantes = "Usuario o contrasena incorrectos."
    End If
End Function

Sub RegistrarFalloLogin(conn, codIndiceUsuario, numFallosActual)
    Dim nuevoNumFallos, umbral, codBloqueado, codBorrado
    nuevoNumFallos = numFallosActual + 1
    Call EjecutarNonQuery(conn, "UPDATE USUARIOS SET NUM_FALLOS = ? WHERE COD_INDICE_USUARIO = ?", Array(nuevoNumFallos, codIndiceUsuario))

    umbral = ValorParametroEntero(conn, "NUM_MAX_INTENTOS_FALLIDOS", 5)
    If nuevoNumFallos >= umbral Then
        codBloqueado = CodigoEstadoUsuario(conn, "Bloqueado")
        codBorrado = CodigoEstadoUsuario(conn, "Borrado")
        Call EjecutarNonQuery(conn, "UPDATE USUARIOS SET ESTADO = ? WHERE COD_INDICE_USUARIO = ? AND ESTADO <> ?", Array(codBloqueado, codIndiceUsuario, codBorrado))
        Call RegistrarAuditoria(conn, codIndiceUsuario, "LOGIN_BLOQUEO_AUTOMATICO", "USUARIOS", codIndiceUsuario, "BLOQUEADO", "Bloqueo automatico tras " & nuevoNumFallos & " intentos fallidos")
    End If
End Sub

' Punto de entrada unico del login. Devuelve un Scripting.Dictionary:
'   Exito (bool); Mensaje (string, solo si Exito=False);
'   si Exito=True ademas: CodIndiceUsuario, CodUsuario, Nombre, Apellidos, CodPerfil, CodOperador, AccionesPermitidas
Function IntentarLogin(conn, codUsuario, passwordPlano)
    Dim resultado
    Set resultado = Server.CreateObject("Scripting.Dictionary")
    resultado("Exito") = False
    resultado("Mensaje") = "Usuario o contrasena incorrectos."

    Dim rs
    Set rs = EjecutarConsulta(conn, _
        "SELECT COD_INDICE_USUARIO, NOMBRE, APELLIDOS, PASSWORD, ALGORITMO_PASSWORD, FECHA_CADUCIDAD, NUM_FALLOS, ESTADO, COD_PERFIL, COD_OPERADOR FROM USUARIOS WHERE COD_USUARIO = ?", _
        Array(codUsuario))

    If rs.EOF Then
        rs.Close
        Call RegistrarAuditoria(conn, Null, "LOGIN", "USUARIOS", Null, "FALLO", "Usuario no encontrado")
        Set IntentarLogin = resultado
        Exit Function
    End If

    Dim codIndiceUsuario, estado, numFallos, hashAlmacenado, algoritmo, fechaCaducidad
    Dim nombre, apellidos, codPerfil, codOperador

    codIndiceUsuario = CLng(rs("COD_INDICE_USUARIO"))
    estado = CLng(rs("ESTADO"))
    If IsNull(rs("NUM_FALLOS")) Then
        numFallos = 0
    Else
        numFallos = CLng(rs("NUM_FALLOS"))
    End If
    hashAlmacenado = rs("PASSWORD") & ""
    algoritmo = rs("ALGORITMO_PASSWORD") & ""
    fechaCaducidad = rs("FECHA_CADUCIDAD")
    nombre = rs("NOMBRE") & ""
    apellidos = rs("APELLIDOS") & ""
    codPerfil = CLng(rs("COD_PERFIL"))
    codOperador = CLng(rs("COD_OPERADOR"))
    rs.Close

    Dim estadoBorrado, estadoBloqueado
    estadoBorrado = CodigoEstadoUsuario(conn, "Borrado")
    estadoBloqueado = CodigoEstadoUsuario(conn, "Bloqueado")

    If estado = estadoBorrado Then
        Call RegistrarAuditoria(conn, codIndiceUsuario, "LOGIN", "USUARIOS", codIndiceUsuario, "FALLO", "Usuario dado de baja")
        Set IntentarLogin = resultado
        Exit Function
    End If

    If estado = estadoBloqueado Then
        resultado("Mensaje") = "Su cuenta ha sido bloqueada por superar el numero de intentos fallidos. Contacte con el administrador de seguridad."
        Call RegistrarAuditoria(conn, codIndiceUsuario, "LOGIN", "USUARIOS", codIndiceUsuario, "FALLO", "Cuenta bloqueada")
        Set IntentarLogin = resultado
        Exit Function
    End If

    If Not VerificarPassword(passwordPlano, hashAlmacenado, algoritmo) Then
        Call RegistrarFalloLogin(conn, codIndiceUsuario, numFallos)
        resultado("Mensaje") = MensajeIntentosRestantes(conn, numFallos + 1)
        Call RegistrarAuditoria(conn, codIndiceUsuario, "LOGIN", "USUARIOS", codIndiceUsuario, "FALLO", "Contrasena incorrecta")
        Set IntentarLogin = resultado
        Exit Function
    End If

    If IsDate(fechaCaducidad) Then
        If CDate(fechaCaducidad) < Now() Then
            resultado("Mensaje") = "Su contrasena ha caducado. Contacte con el administrador para renovarla."
            Call RegistrarAuditoria(conn, codIndiceUsuario, "LOGIN", "USUARIOS", codIndiceUsuario, "FALLO", "Password caducada")
            Set IntentarLogin = resultado
            Exit Function
        End If
    End If

    Call RegistrarLoginCorrecto(conn, codIndiceUsuario)

    Dim debeRehash, partesHash
    debeRehash = (algoritmo <> "PBKDF2SHA256")
    If Not debeRehash Then
        partesHash = Split(hashAlmacenado, "$")
        If UBound(partesHash) >= 1 And IsNumeric(partesHash(1)) Then
            debeRehash = (CLng(partesHash(1)) <> CLng(Application("PBKDF2Iteraciones")))
        End If
    End If
    If debeRehash Then
        Call RehashPassword(conn, codIndiceUsuario, passwordPlano)
    End If

    resultado("Exito") = True
    resultado("Mensaje") = ""
    resultado("CodIndiceUsuario") = codIndiceUsuario
    resultado("CodUsuario") = codUsuario
    resultado("Nombre") = nombre
    resultado("Apellidos") = apellidos
    resultado("CodPerfil") = codPerfil
    resultado("CodOperador") = codOperador
    resultado("AccionesPermitidas") = AccionesPermitidas(conn, codPerfil)

    Call RegistrarAuditoria(conn, codIndiceUsuario, "LOGIN", "USUARIOS", codIndiceUsuario, "OK", "")

    Set IntentarLogin = resultado
End Function
%>
