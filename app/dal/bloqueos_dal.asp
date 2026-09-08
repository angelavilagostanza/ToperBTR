<%
' DAL de Bloqueos — completo (Entregable 6).
' Las funciones BuscarBloqueoActivoPorImei y ExcluirIMEI ya existian desde el Entregable 3
' (las necesitaba solicitudes/alta_do.asp para excluir un IMEI al crear una solicitud de
' Exclusion). Se mantienen sin cambios; el resto es nuevo.

' Usan EjecutarConsultaTx / EjecutarNonQueryTx (variantes sin captura automatica de errores)
' porque todas se invocan desde funciones que ya gestionan su propia transaccion
' (InsertarBloqueoDirecto, DesbloquearIMEIsLote) o porque se llaman a su vez desde
' transacciones ajenas (ExcluirIMEI desde CrearSolicitudCompleta). Ver conexion.asp.

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

' tipoAExcluirDescripcion: "BLOQUEO DIRECTO" o "BLOQUEO GLOBAL".
' Corrección deliberada: TODAS las ramas historizan en HISTORICO_BLOQUEO_DIRECTO de forma
' simetrica (el legacy solo historizaba al excluir la componente "Directo", no la "Global").
' NO abre ni cierra transaccion propia — el llamador decide el alcance.
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
	
	If codTipoAExcluir = codDirecto Then
		If huboExclusion And Err.Number = 0 Then
			Call EjecutarNonQueryTx(conn, _
				"INSERT INTO HISTORICO_BLOQUEO_DIRECTO (IMEI, FECHA_INCLUSION, FECHA_EXCLUSION, COD_RAZON, EN_EIR) VALUES (?, ?, NOW(), ?, 1)", _
				Array(bloqueo("Imei"), bloqueo("FechaInclusion"), bloqueo("CodRazon")))
		End If
	End If

    ExcluirIMEI = huboExclusion And (Err.Number = 0)
End Function

' Comprueba si el IMEI pertenece a un cliente Yoigo (existe en LISTA_BLANCA).
' Regla de negocio confirmada en DAOEIR.insertarBloqueo / buscarIMEIduplicado: el bloqueo
' directo manual solo se permite para IMEIs reconocidos como clientes Yoigo.
' La busqueda usa los 14 primeros digitos del IMEI (igual que el legacy), ya que el digito
' 15 es el digito de control Luhn, que puede variar entre variantes del mismo dispositivo.
' NOTA a verificar: el legacy usaba ademas un filtro de dias (parametro DIAS_MARGEN de
' PARAMETROS_REF) sobre una columna de fecha de LISTA_BLANCA. Sin conocer el nombre exacto
' de esa columna en la BD real, aqui solo se comprueba existencia. Si se necesita el filtro
' de dias, anadir: AND <columna_fecha> >= DATE_SUB(NOW(), INTERVAL ? DAY)
Function EsIMEICliente(conn, imei)
    Dim rs, prefijo
    prefijo = Left(imei, 14)
    Set rs = EjecutarConsultaTx(conn, _
        "SELECT 1 FROM LISTA_BLANCA WHERE IMEI LIKE ?", _
        Array(prefijo & "%"))
    EsIMEICliente = Not rs.EOF
    rs.Close
End Function

' Reproduce la logica de DAOEIR.insertarBloqueo con dos correcciones:
'   1. Todo el flujo va dentro de una transaccion (el legacy no la usaba — riesgo TOCTOU).
'   2. Consultas 100% parametrizadas (el legacy concatenaba valores directamente).
'
' Retorna: "OK"        — bloqueo insertado/actualizado correctamente.
'          "BLOQUEADO" — el IMEI ya estaba bloqueado con el mismo tipo o con COMBINADO.
'          "NO_YOIGO"  — el IMEI no existe en LISTA_BLANCA (no es cliente Yoigo).
'          "ERROR"     — fallo de BD; la transaccion fue revertida.
'
' codRazon: codigo numerico de RAZON_LISTA_NEGRA_REF.
' comentario: puede ser vacio.
Function InsertarBloqueoDirecto(conn, imei, codRazon, comentario)
    On Error Resume Next

    Dim bloqueo, codDirecto, codCombinado, resultado
    resultado = "ERROR"

    If Not EsIMEICliente(conn, imei) Then
        InsertarBloqueoDirecto = "NO_YOIGO"
        Exit Function
    End If

    conn.BeginTrans

    codDirecto = CodigoTipoBloqueo(conn, "BLOQUEO DIRECTO")
    codCombinado = CodigoTipoBloqueo(conn, "BLOQUEO DIRECTO Y GLOBAL")
    If Err.Number <> 0 Then conn.RollbackTrans : InsertarBloqueoDirecto = "ERROR" : Exit Function

    Set bloqueo = BuscarBloqueoActivoPorImei(conn, imei)
    If Err.Number <> 0 Then conn.RollbackTrans : InsertarBloqueoDirecto = "ERROR" : Exit Function

    If Not (bloqueo Is Nothing) Then
        ' IMEI ya en LISTA_NEGRA
        If bloqueo("CodTipoBloqueo") = codDirecto Or bloqueo("CodTipoBloqueo") = codCombinado Then
            conn.RollbackTrans
            InsertarBloqueoDirecto = "BLOQUEADO"
            Exit Function
        End If
        ' Esta en LISTA_NEGRA con BLOQUEO GLOBAL -> ampliar a COMBINADO
        Call EjecutarNonQueryTx(conn, _
            "UPDATE LISTA_NEGRA SET COD_TIPO_BLOQUEO = ? WHERE COD_LISTA_NEGRA = ?", _
            Array(codCombinado, bloqueo("CodListaNegra")))
        If Err.Number <> 0 Then conn.RollbackTrans : InsertarBloqueoDirecto = "ERROR" : Exit Function
        resultado = "OK"
    Else
        ' IMEI no esta en LISTA_NEGRA — comprobar si hay un desbloqueo pendiente
        ' en HISTORICO_BLOQUEO_DIRECTO (EN_EIR=1: notificacion pendiente de enviar al EIR).
        ' Si existe, cancelamos esa notificacion (ponemos EN_EIR=0 en el historico) e
        ' insertamos el nuevo bloqueo con EN_EIR=0 (el EIR nunca supo del bloqueo+desbloqueo).
        ' Si no existe, insertamos con EN_EIR=1 para que el proximo fichero EIR lo incluya.
        Dim hayPendiente, enEir
        Dim rsPend
        Set rsPend = EjecutarConsultaTx(conn, _
            "SELECT 1 FROM HISTORICO_BLOQUEO_DIRECTO WHERE IMEI LIKE ? AND EN_EIR = 1", _
            Array(Left(imei, 14) & "%"))
        If Err.Number <> 0 Then conn.RollbackTrans : InsertarBloqueoDirecto = "ERROR" : Exit Function
        hayPendiente = Not rsPend.EOF
        rsPend.Close

        If hayPendiente Then
            Call EjecutarNonQueryTx(conn, _
                "UPDATE HISTORICO_BLOQUEO_DIRECTO SET EN_EIR = 0 WHERE IMEI LIKE ? AND EN_EIR = 1", _
                Array(Left(imei, 14) & "%"))
            If Err.Number <> 0 Then conn.RollbackTrans : InsertarBloqueoDirecto = "ERROR" : Exit Function
            enEir = 0
        Else
            enEir = 1
        End If

        Call EjecutarNonQueryTx(conn, _
            "INSERT INTO LISTA_NEGRA (IMEI, COD_TIPO_BLOQUEO, COD_RAZON, EN_EIR, FECHA_INCLUSION, COMENTARIO) " & _
            "VALUES (?, ?, ?, ?, NOW(), ?)", _
            Array(imei, codDirecto, codRazon, enEir, comentario))
        If Err.Number <> 0 Then conn.RollbackTrans : InsertarBloqueoDirecto = "ERROR" : Exit Function
        resultado = "OK"
    End If

    conn.CommitTrans
    InsertarBloqueoDirecto = resultado
End Function

' Listado de bloqueos activos con filtros opcionales, todos parametrizados.
' Corrige la inyeccion SQL confirmada en DAOEIR.listarIMEIs (el legacy concatenaba
' IMEI, razon, fechas y orden directamente en el SQL sin PreparedStatement real).
' Cuando se filtra por tipo, incluye tambien las filas BLOQUEO DIRECTO Y GLOBAL (que
' contienen la componente del tipo filtrado), igual que hacia el legacy, pero resolviendo
' el codigo del tipo COMBINADO por catalogo en vez de hardcodearlo como hacia el Java.
' Devuelve un Recordset — cerrar en el llamador.
Function ListarBloqueos(conn, imei, codTipo, codRazon, fechaDesde, fechaHasta)
    Dim sql, paramsList(), paramCount
    paramCount = 0
    ReDim paramsList(5)

    sql = "SELECT LN.COD_LISTA_NEGRA, LN.IMEI, LN.FECHA_INCLUSION, " & _
          "TB.DESCRIPCION AS TIPO_BLOQUEO_DESC, LN.COD_RAZON, " & _
          "RLN.DESCRIPCION AS RAZON_DESC, LN.COMENTARIO, LN.EN_EIR " & _
          "FROM LISTA_NEGRA LN " & _
          "LEFT JOIN TIPO_BLOQUEO_REF TB ON TB.COD_TIPO_BLOQUEO = LN.COD_TIPO_BLOQUEO " & _
          "LEFT JOIN RAZON_LISTA_NEGRA_REF RLN ON RLN.COD_RAZON = LN.COD_RAZON " & _
          "WHERE 1=1"

    If Len(imei) >= 14 Then
        sql = sql & " AND LN.IMEI LIKE ?"
        paramsList(paramCount) = Left(imei, 14) & "%"
        paramCount = paramCount + 1
    End If

    If CLng(codTipo & "0") > 0 Then
        Dim codCombinado
        codCombinado = CodigoTipoBloqueo(conn, "BLOQUEO DIRECTO Y GLOBAL")
        sql = sql & " AND LN.COD_TIPO_BLOQUEO IN (?, ?)"
        paramsList(paramCount) = CLng(codTipo)
        paramCount = paramCount + 1
        paramsList(paramCount) = codCombinado
        paramCount = paramCount + 1
    End If

    If CLng(codRazon & "0") > 0 Then
        sql = sql & " AND LN.COD_RAZON = ?"
        paramsList(paramCount) = CLng(codRazon)
        paramCount = paramCount + 1
    End If

    If Trim(fechaDesde & "") <> "" Then
        sql = sql & " AND LN.FECHA_INCLUSION >= ?"
        paramsList(paramCount) = CDate(fechaDesde)
        paramCount = paramCount + 1
    End If

    If Trim(fechaHasta & "") <> "" Then
        sql = sql & " AND LN.FECHA_INCLUSION < DATE_ADD(?, INTERVAL 1 DAY)"
        paramsList(paramCount) = CDate(fechaHasta)
        paramCount = paramCount + 1
    End If

    sql = sql & " ORDER BY LN.FECHA_INCLUSION DESC"

    If paramCount = 0 Then
        Set ListarBloqueos = EjecutarConsulta(conn, sql, Array())
    Else
        ReDim Preserve paramsList(paramCount - 1)
        Set ListarBloqueos = EjecutarConsulta(conn, sql, paramsList)
    End If
End Function

' Desbloquea un array de IMEIs (componente BLOQUEO DIRECTO), cada uno en su propia
' transaccion independiente para que un fallo en uno no revierta los demas.
' Retorna el numero de IMEIs desbloqueados correctamente.
' Llama a ExcluirIMEI, que NO abre transaccion propia — esta funcion la abre/cierra.
Function DesbloquearIMEIsLote(conn, imeisArray)
    Dim i, ok, desbloqueados
    desbloqueados = 0
    For i = 0 To UBound(imeisArray)
        Dim imei
        imei = Trim(imeisArray(i))
        If Len(imei) = 15 Then
            On Error Resume Next
            conn.BeginTrans
            ok = ExcluirIMEI(conn, imei, "BLOQUEO DIRECTO")
            If Err.Number <> 0 Then
                conn.RollbackTrans
            ElseIf ok Then
                conn.CommitTrans
                desbloqueados = desbloqueados + 1
            Else
                conn.RollbackTrans
            End If
            On Error Goto 0
        End If
    Next
    DesbloquearIMEIsLote = desbloqueados
End Function
%>
