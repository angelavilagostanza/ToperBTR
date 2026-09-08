<%
' DAL de Ficheros (Entregable 5): busqueda/listado/descarga de los ficheros de
' intercambio y consolidacion que genera EIRControl (su generacion queda fuera de
' alcance, este modulo solo los consulta). El legacy real (DAOFicheros.buscarFicheros)
' concatenaba TODOS los filtros - nombre, fechas, operador - directamente en el SQL sin
' parametrizar, confirmado como inyeccion SQL real (ver memoria toperbtr-java-findings).
' Aqui todo va parametrizado.
'
' El "tipo" de fichero (Diario/Consolidacion/EIR) y el operador propietario NO son
' columnas reales de FICHEROS: se infieren de la longitud y del primer caracter del
' NOMBRE, tal y como hace el codigo legacy real (FicherosBusqVO.getTipoFichero()/
' esTipoEIR() y DAOFicheros.buscarFicheros). Es una regla de negocio fragil basada en el
' nombre de fichero, no una decision de diseno nueva de ToperBTR - se preserva porque asi
' es como EIRControl nombra los ficheros hoy y no se puede cambiar sin tocar EIRControl
' (fuera de alcance).
Function BuscarFicheros(conn, filtroNombre, filtroTipo, filtroOperador, filtroFechaDesde, filtroFechaHasta, esAdministrador)
    Dim sql, listaParams(), n, codFicheroGenerico
    codFicheroGenerico = ValorParametroEntero(conn, "COD_FICHERO_GENERICO", 0)

    sql = "SELECT F.COD_FICHERO, F.NOMBRE, F.FECHA_CREACION, " & _
          "(SELECT COUNT(*) FROM INCIDENCIA_TO_OBJETO ITO WHERE ITO.TIPO_OBJETO = 'F' AND ITO.COD_OBJETO = F.COD_FICHERO) AS NUM_INCIDENCIAS " & _
          "FROM FICHEROS F WHERE F.COD_FICHERO <> ?"
    ReDim listaParams(0)
    listaParams(0) = codFicheroGenerico
    n = 1

    If EsCadenaNoVacia(filtroNombre) Then
        sql = sql & " AND F.NOMBRE = ?"
        ReDim Preserve listaParams(n) : listaParams(n) = UCase(filtroNombre) : n = n + 1
    End If

    If IsDate(filtroFechaDesde) Then
        sql = sql & " AND F.FECHA_CREACION >= ?"
        ReDim Preserve listaParams(n) : listaParams(n) = CDate(filtroFechaDesde) : n = n + 1
    End If
    If IsDate(filtroFechaHasta) Then
        sql = sql & " AND F.FECHA_CREACION < ?"
        ReDim Preserve listaParams(n) : listaParams(n) = DateAdd("d", 1, CDate(filtroFechaHasta)) : n = n + 1
    End If

    If EsCadenaNoVacia(filtroOperador) And filtroOperador <> "-1" And IsNumeric(filtroOperador) Then
        Dim rsOp, claveOp, nombreOp
        claveOp = ""
        Set rsOp = EjecutarConsultaTx(conn, "SELECT NOMBRE, CLAVE FROM OPERADOR_REF WHERE COD_OPERADOR = ?", Array(CLng(filtroOperador)))
        If Not rsOp.EOF Then
            nombreOp = rsOp("NOMBRE") & ""
            claveOp = rsOp("CLAVE") & ""
        End If
        rsOp.Close
        If EsCadenaNoVacia(claveOp) Then
            ' Yoigo / Grupo MASMOVIL / GMM son el mismo operador — cualquiera de los tres
            ' nombres puede aparecer en OPERADOR_REF.NOMBRE segun el estado de la BD.
            Dim esGMMOp
            esGMMOp = (StrComp(nombreOp, "Yoigo",          vbTextCompare) = 0 Or _
                       StrComp(nombreOp, "Grupo MASMOVIL",  vbTextCompare) = 0 Or _
                       StrComp(nombreOp, "GMM",             vbTextCompare) = 0)
            If esGMMOp Then
                ' Regla real del legacy: los ficheros del feed EIR de este operador no
                ' llevan su clave en el nombre, asi que buscar por su operador tambien
                ' debe traerlos (coinciden por contener "EIR" en vez de la clave).
                sql = sql & " AND (F.NOMBRE LIKE ? OR F.NOMBRE LIKE '%EIR%')"
            Else
                sql = sql & " AND F.NOMBRE LIKE ?"
            End If
            ReDim Preserve listaParams(n) : listaParams(n) = "%" & claveOp & "%" : n = n + 1
        End If
    End If

    ' EIR (longitud 11) solo visible para el perfil Administrador - gate real en
    ' servidor, no solo ocultando la opcion en el formulario de busqueda.
    If Not esAdministrador Then
        sql = sql & " AND LENGTH(F.NOMBRE) <> 11"
    End If

    Select Case UCase(filtroTipo & "")
        Case "D"
            sql = sql & " AND LENGTH(F.NOMBRE) = 9"
        Case "C"
            sql = sql & " AND LENGTH(F.NOMBRE) = 7"
        Case "E"
            sql = sql & " AND LENGTH(F.NOMBRE) = 11"
    End Select

    sql = sql & " ORDER BY F.FECHA_CREACION DESC"
    Set BuscarFicheros = EjecutarConsulta(conn, sql, listaParams)
End Function

' Usada tanto por la descarga como por cualquier lectura puntual de un fichero por su
' nombre. Devuelve Nothing si no existe - la ruta fisica viene siempre de la BD, nunca
' se construye a partir de lo que escriba el usuario (evita cualquier traversal de rutas).
Function ObtenerFicheroPorNombre(conn, nombre)
    Dim rs, f
    Set rs = EjecutarConsulta(conn, _
        "SELECT COD_FICHERO, NOMBRE, EXTENSION, RUTA, FECHA_CREACION FROM FICHEROS WHERE NOMBRE = ?", _
        Array(nombre))
    If rs.EOF Then
        Set ObtenerFicheroPorNombre = Nothing
    Else
        Set f = Server.CreateObject("Scripting.Dictionary")
        f.Add "CodFichero", CLng(rs("COD_FICHERO"))
        f.Add "NombreFichero", rs("NOMBRE") & ""
        f.Add "Extension", rs("EXTENSION") & ""
        f.Add "Ruta", rs("RUTA") & ""
        f.Add "FechaCreacion", rs("FECHA_CREACION")
        Set ObtenerFicheroPorNombre = f
    End If
    rs.Close
End Function
%>
