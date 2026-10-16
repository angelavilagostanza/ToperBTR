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

' Antes esto recibia numFallosActual (leido al PRINCIPIO de la peticion, antes de
' verificar la contrasena) y escribia numFallosActual+1 calculado en VBScript. Bajo
' peticiones en paralelo (p.ej. un ataque de fuerza bruta con varios hilos), todas leian
' el mismo NUM_FALLOS inicial y todas escribian el mismo "+1", asi que el contador real
' nunca avanzaba al ritmo de los intentos reales - el bloqueo por intentos fallidos se
' podia evitar por completo con peticiones simultaneas (hallazgo A1). Ahora el incremento
' es atomico en SQL (NUM_FALLOS = NUM_FALLOS + 1, no un valor precalculado), con el
' bloqueo de fila que MySQL aplica durante el UPDATE, y se relee el valor ya actualizado
' para decidir si toca bloquear la cuenta - nunca se trabaja con un valor que pudo
' quedarse desfasado por otra peticion concurrente.
Function RegistrarFalloLogin(conn, codIndiceUsuario)
    Dim nuevoNumFallos, umbral, codBloqueado, codBorrado, rs

    ' COALESCE(NUM_FALLOS, 0): algunas cuentas tienen NUM_FALLOS a NULL en BD (nunca han
    ' fallado un login desde que existe la columna - el propio IntentarLogin de mas abajo
    ' ya comprobaba IsNull al leerlo). "NULL + 1" en SQL se queda en NULL para siempre, lo
    ' que ademas provocaba un error VBScript sin capturar ("Invalid use of Null") al hacer
    ' CLng(Null) mas abajo - confirmado en pruebas en pre (pagina en blanco, sin redirigir).
    Call EjecutarNonQuery(conn, "UPDATE USUARIOS SET NUM_FALLOS = COALESCE(NUM_FALLOS, 0) + 1 WHERE COD_INDICE_USUARIO = ?", Array(codIndiceUsuario))

    Set rs = EjecutarConsulta(conn, "SELECT NUM_FALLOS FROM USUARIOS WHERE COD_INDICE_USUARIO = ?", Array(codIndiceUsuario))
    If rs.EOF Then
        nuevoNumFallos = 1
    ElseIf IsNull(rs("NUM_FALLOS")) Then
        nuevoNumFallos = 1
    Else
        nuevoNumFallos = CLng(rs("NUM_FALLOS"))
    End If
    rs.Close

    umbral = ValorParametroEntero(conn, "NUM_MAX_INTENTOS_FALLIDOS", 5)
    If nuevoNumFallos >= umbral Then
        codBloqueado = CodigoEstadoUsuario(conn, "Bloqueado")
        codBorrado = CodigoEstadoUsuario(conn, "Borrado")
        Call EjecutarNonQuery(conn, "UPDATE USUARIOS SET ESTADO = ? WHERE COD_INDICE_USUARIO = ? AND ESTADO <> ?", Array(codBloqueado, codIndiceUsuario, codBorrado))
        Call RegistrarAuditoria(conn, codIndiceUsuario, "LOGIN_BLOQUEO_AUTOMATICO", "USUARIOS", codIndiceUsuario, "BLOQUEADO", "Bloqueo automatico tras " & nuevoNumFallos & " intentos fallidos")
    End If

    RegistrarFalloLogin = nuevoNumFallos
End Function

' --- Limite de intentos de login por IP (hallazgo A2) ---------------------------------
' El bloqueo de arriba es POR CUENTA; no impide que alguien pruebe muchas cuentas
' distintas (o la misma, con contrasenas distintas) desde la misma IP sin disparar nunca
' el bloqueo de ninguna cuenta en concreto. Esto es un segundo control independiente, en
' memoria de aplicacion (sin tocar BD - se reinicia solo si el Application Pool recicla,
' lo cual es aceptable para este proposito).
'
' IMPORTANTE: el primer intento de esto guardaba un Scripting.Dictionary en Application,
' y eso tumbo TODA la aplicacion de inmediato (no solo el login) - Scripting.Dictionary es
' un objeto de apartamento unico (STA) y NO es seguro guardarlo en Application, que IIS
' comparte entre peticiones concurrentes potencialmente en hilos distintos; es un problema
' clasico y documentado de ASP clasico. Por eso aqui NO se guarda ningun objeto: cada IP
' tiene su propia variable de Application con un valor de texto plano "contador|inicio"
' (string, tipo de dato simple sin ningun problema de apartamento). "inicio" se guarda
' como numero (CDbl de la fecha, no CStr/CDate de la fecha formateada) para no depender
' del locale regional del servidor al convertir de ida y vuelta.
Const RATE_LIMIT_IP_VENTANA_MINUTOS = 5
Const RATE_LIMIT_IP_MAX_INTENTOS = 20

Function ClaveAppIP(ip)
    ClaveAppIP = "RL_IP_" & Replace(Replace(ip, ".", "_"), ":", "_")
End Function

Function PermiteIntentoLoginIP(ip)
    Dim clave, valor, partes
    PermiteIntentoLoginIP = True
    If Not EsCadenaNoVacia(ip) Then Exit Function

    clave = ClaveAppIP(ip)
    Application.Lock
    valor = "" & Application(clave)
    Application.Unlock

    If valor <> "" Then
        partes = Split(valor, "|")
        If UBound(partes) = 1 Then
            If DateDiff("n", CDate(CDbl(partes(1))), Now()) < RATE_LIMIT_IP_VENTANA_MINUTOS Then
                If CLng(partes(0)) >= RATE_LIMIT_IP_MAX_INTENTOS Then
                    PermiteIntentoLoginIP = False
                End If
            End If
        End If
    End If
End Function

Sub RegistrarFalloLoginIP(ip)
    Dim clave, valor, partes, contador, inicio, parseable
    If Not EsCadenaNoVacia(ip) Then Exit Sub

    clave = ClaveAppIP(ip)
    parseable = False

    Application.Lock
    valor = "" & Application(clave)
    If valor <> "" Then
        partes = Split(valor, "|")
        If UBound(partes) = 1 Then
            parseable = True
            contador = CLng(partes(0))
            inicio = CDate(CDbl(partes(1)))
        End If
    End If

    If parseable Then
        If DateDiff("n", inicio, Now()) >= RATE_LIMIT_IP_VENTANA_MINUTOS Then
            contador = 1
            inicio = Now()
        Else
            contador = contador + 1
        End If
    Else
        contador = 1
        inicio = Now()
    End If

    Application(clave) = contador & "|" & CStr(CDbl(inicio))
    Application.Unlock
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

    Dim codIndiceUsuario, estado, hashAlmacenado, algoritmo, fechaCaducidad
    Dim nombre, apellidos, codPerfil, codOperador

    ' NUM_FALLOS ya no se lee aqui: RegistrarFalloLogin (abajo) hace su propio incremento
    ' atomico y relee el valor actualizado, precisamente para no depender de un valor que
    ' pudo quedarse desfasado por otra peticion concurrente (hallazgo A1).
    codIndiceUsuario = CLng(rs("COD_INDICE_USUARIO"))
    estado = CLng(rs("ESTADO"))
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
        Dim nuevoNumFallos
        nuevoNumFallos = RegistrarFalloLogin(conn, codIndiceUsuario)
        resultado("Mensaje") = MensajeIntentosRestantes(conn, nuevoNumFallos)
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
    resultado("NombrePerfil") = NombrePerfil(conn, codPerfil)
    resultado("CodOperador") = codOperador
    resultado("AccionesPermitidas") = AccionesPermitidas(conn, codPerfil)

    Call RegistrarAuditoria(conn, codIndiceUsuario, "LOGIN", "USUARIOS", codIndiceUsuario, "OK", "")

    Set IntentarLogin = resultado
End Function
%>
