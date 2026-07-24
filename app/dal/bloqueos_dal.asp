<%
' Version minima del DAL de Bloqueos: solo lo que necesita Solicitudes (Entregable 3)
' para el efecto colateral de excluir un IMEI de LISTA_NEGRA cuando se crea una
' solicitud de Exclusion. El resto del modulo (bloqueo directo manual, consulta,
' pantallas propias) se completa en el Entregable 6, reutilizando este mismo fichero.

' Se llama siempre desde ExcluirIMEI, que a su vez se llama siempre dentro de una
' transaccion ya abierta por el llamador (CrearSolicitudCompleta, o en el futuro
' Entregable 6) - usa la variante Tx (ver nota en include/data/conexion.asp).
Function BuscarBloqueoActivoPorImei(conn, imei)
    Dim rs, b, prefijo
    prefijo = Left(imei, 14)
    Set rs = EjecutarConsultaTx(conn, _
        "SELECT COD_LISTA_NEGRA, IMEI, FECHA_INCLUSION, COD_TIPO_BLOQUEO, COD_RAZON, EN_EIR FROM LISTA_NEGRA WHERE IMEI LIKE ?", _
        Array(prefijo & "%"))
    If rs.EOF Then
        Set BuscarBloqueoActivoPorImei = Nothing
    Else
        Set b = Server.CreateObject("Scripting.Dictionary")
        b("CodListaNegra") = CLng(rs("COD_LISTA_NEGRA"))
        b("Imei") = rs("IMEI") & ""
        b("FechaInclusion") = rs("FECHA_INCLUSION")
        b("CodTipoBloqueo") = CLng(rs("COD_TIPO_BLOQUEO"))
        b("CodRazon") = CLng(rs("COD_RAZON"))
        b("EnEir") = CLng(rs("EN_EIR"))
        Set BuscarBloqueoActivoPorImei = b
    End If
    rs.Close
End Function

' Reproduce la logica confirmada de DAOEIR.excluyeIMEI del legacy, con una correccion
' deliberada: TODAS las ramas historizan en HISTORICO_BLOQUEO_DIRECTO de forma simetrica
' (decision de negocio del usuario - el legacy solo historizaba al excluir el componente
' "Directo" de un bloqueo, nunca al excluir el componente "Global", asimetria real
' confirmada en el codigo fuente). Simplificacion asumida (a verificar si hiciera falta
' mas precision): EN_EIR de LISTA_NEGRA solo se toca en las ramas de tipo exacto (fiel al
' legacy); en las ramas de bloqueo combinado no se modifica EN_EIR de LISTA_NEGRA (el
' registro sigue activo, solo cambia su tipo), y el HISTORICO siempre usa EN_EIR=1
' ("pendiente de notificar en el proximo fichero EIR") en cualquier exclusion.
'
' tipoAExcluirDescripcion: "BLOQUEO DIRECTO" o "BLOQUEO GLOBAL".
' Devuelve True si habia un bloqueo activo y se ha excluido; False si no habia nada que
' excluir. IMPORTANTE: esta funcion NO abre ni cierra su propia transaccion - ADO no
' soporta transacciones anidadas en la misma conexion, y esta funcion se reutiliza tanto
' desde CrearSolicitudCompleta (como parte de una transaccion mayor) como, en el futuro
' Entregable 6, desde una pantalla de bloqueos con su propia transaccion. El llamador es
' quien decide el alcance de BeginTrans/CommitTrans/RollbackTrans; si algo falla aqui,
' Err.Number queda distinto de 0 (el objeto Err es global) para que el llamador lo detecte
' justo despues de invocar esta funcion, igual que con cualquier otra sentencia suya.
Function ExcluirIMEI(conn, imei, tipoAExcluirDescripcion)
    On Error Resume Next

    Dim bloqueo, codTipoAExcluir, codDirecto, codGlobal, codCombinado, huboExclusion
    huboExclusion = False

    Set bloqueo = BuscarBloqueoActivoPorImei(conn, imei)
    If bloqueo Is Nothing Then
        ExcluirIMEI = False
        Exit Function
    End If

    codDirecto = CodigoTipoBloqueo(conn, "BLOQUEO DIRECTO")
    codGlobal = CodigoTipoBloqueo(conn, "BLOQUEO GLOBAL")
    codCombinado = CodigoTipoBloqueo(conn, "BLOQUEO DIRECTO Y GLOBAL")
    codTipoAExcluir = CodigoTipoBloqueo(conn, tipoAExcluirDescripcion)

    If bloqueo("CodTipoBloqueo") = codTipoAExcluir Then
        If codTipoAExcluir = codDirecto Then
            Call EjecutarNonQueryTx(conn, "DELETE FROM LISTA_NEGRA WHERE COD_LISTA_NEGRA = ?", Array(bloqueo("CodListaNegra")))
        Else
            Call EjecutarNonQueryTx(conn, "UPDATE LISTA_NEGRA SET EN_EIR = 2 WHERE COD_LISTA_NEGRA = ?", Array(bloqueo("CodListaNegra")))
        End If
        huboExclusion = True
    ElseIf bloqueo("CodTipoBloqueo") = codCombinado Then
        If codTipoAExcluir = codDirecto Then
            Call EjecutarNonQueryTx(conn, "UPDATE LISTA_NEGRA SET COD_TIPO_BLOQUEO = ? WHERE COD_LISTA_NEGRA = ?", Array(codGlobal, bloqueo("CodListaNegra")))
        Else
            Call EjecutarNonQueryTx(conn, "UPDATE LISTA_NEGRA SET COD_TIPO_BLOQUEO = ? WHERE COD_LISTA_NEGRA = ?", Array(codDirecto, bloqueo("CodListaNegra")))
        End If
        huboExclusion = True
    End If

    If huboExclusion And Err.Number = 0 Then
        Call EjecutarNonQueryTx(conn, _
            "INSERT INTO HISTORICO_BLOQUEO_DIRECTO (IMEI, FECHA_INCLUSION, FECHA_EXCLUSION, COD_RAZON, EN_EIR) VALUES (?, ?, NOW(), ?, 1)", _
            Array(bloqueo("Imei"), bloqueo("FechaInclusion"), bloqueo("CodRazon")))
    End If

    ExcluirIMEI = huboExclusion And (Err.Number = 0)
End Function
%>
