<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/solicitudes_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequiereAccion(conn, CodigoAccion(conn, "LECTURA_SOLICITUDES"))

Dim puedeEscribir
puedeEscribir = TienePermiso(Session("AccionesPermitidas"), CodigoAccion(conn, "ESCRITURA_SOLICITUDES"))

Dim token
token = GenerarTokenCsrf()

Dim fCod, fImei, fMsisdn, fTipo, fEstado, fIdent, fNombre, fDesde, fHasta
Dim buscado, limite, pagina, totalRegistros, totalPaginas, offsetReg

fCod    = Trim(Request.QueryString("codSolicitud"))
fImei   = Trim(Request.QueryString("imei"))
fMsisdn = Trim(Request.QueryString("msisdn"))
fTipo   = Request.QueryString("tipo")
fEstado = Request.QueryString("estado")
fIdent  = Trim(Request.QueryString("identCliente"))
fNombre = Trim(Request.QueryString("nombreCliente"))
fDesde  = Trim(Request.QueryString("fechaDesde"))
fHasta  = Trim(Request.QueryString("fechaHasta"))
buscado = (Request.QueryString("buscado") = "1")

Select Case CStr(Request.QueryString("limite"))
    Case "100":  limite = 100
    Case "1000": limite = 1000
    Case Else:   limite = 25
End Select

pagina = 1
If IsNumeric(Request.QueryString("pagina")) Then
    If CLng(Request.QueryString("pagina")) > 1 Then
        pagina = CLng(Request.QueryString("pagina"))
    End If
End If

' Cargar catalogos en arrays y cerrar sus cursores de inmediato.
' MySQL ODBC 8.0 no soporta multiples cursores server-side abiertos simultaneamente
' sobre la misma conexion; cerrarlos antes de abrir el cursor principal evita que
' los selects de catalogos corrompan los datos del cursor principal.
Dim rsTmp
Dim arrTipoCod(), arrTipoDesc(), nTipos
nTipos = 0
Set rsTmp = ListarTiposSolicitud(conn)
Do While Not rsTmp.EOF
    ReDim Preserve arrTipoCod(nTipos)
    ReDim Preserve arrTipoDesc(nTipos)
    arrTipoCod(nTipos)  = rsTmp("COD_TIPO_SOLICITUD")
    arrTipoDesc(nTipos) = rsTmp("DESCRIPCION") & ""
    nTipos = nTipos + 1
    rsTmp.MoveNext
Loop
rsTmp.Close

Dim arrEstadoCod(), arrEstadoDesc(), nEstados
nEstados = 0
Set rsTmp = ListarEstadosSolicitud(conn)
Do While Not rsTmp.EOF
    ReDim Preserve arrEstadoCod(nEstados)
    ReDim Preserve arrEstadoDesc(nEstados)
    arrEstadoCod(nEstados)  = rsTmp("COD_ESTADO")
    arrEstadoDesc(nEstados) = rsTmp("DESCRIPCION") & ""
    nEstados = nEstados + 1
    rsTmp.MoveNext
Loop
rsTmp.Close
Set rsTmp = Nothing

' Con los catalogos cerrados, ejecutar la busqueda y pre-cargar todas las filas
' en arrays locales antes de cerrar el cursor. Asi EsSolicitudCancelable puede
' abrir sus propias consultas sin interferir con el cursor de solicitudes.
Dim aIndice(), aCodSol(), aTipo(), aImei(), aMsisdn(), aCliente(), aEstado(), aFecha()
Dim nFilas
nFilas = 0

If buscado Then
    totalRegistros = ContarSolicitudes(conn, fCod, fImei, fMsisdn, fTipo, fEstado, fIdent, fNombre, fDesde, fHasta)
    totalPaginas   = Int((totalRegistros + limite - 1) / limite)
    If totalPaginas < 1 Then totalPaginas = 1
    If pagina > totalPaginas Then pagina = totalPaginas
    offsetReg      = (pagina - 1) * limite

    Dim rsSol
    Set rsSol = BuscarSolicitudes(conn, fCod, fImei, fMsisdn, fTipo, fEstado, fIdent, fNombre, fDesde, fHasta, limite, offsetReg)
    Do While Not rsSol.EOF
        ReDim Preserve aIndice(nFilas)
        ReDim Preserve aCodSol(nFilas)
        ReDim Preserve aTipo(nFilas)
        ReDim Preserve aImei(nFilas)
        ReDim Preserve aMsisdn(nFilas)
        ReDim Preserve aCliente(nFilas)
        ReDim Preserve aEstado(nFilas)
        ReDim Preserve aFecha(nFilas)
        aIndice(nFilas)  = CLng(rsSol("INDICE_SOLICITUD") & "")
        aCodSol(nFilas)  = rsSol("COD_SOLICITUD") & ""
        aTipo(nFilas)    = rsSol("TIPO") & ""
        aImei(nFilas)    = rsSol("IMEI") & ""
        aMsisdn(nFilas)  = rsSol("MSISDN") & ""
        aCliente(nFilas) = rsSol("CLIENTE") & ""
        aEstado(nFilas)  = rsSol("ESTADO_VIGENTE") & ""
        aFecha(nFilas)   = rsSol("FECHA_CREACION")
        nFilas = nFilas + 1
        rsSol.MoveNext
    Loop
    rsSol.Close
    Set rsSol = Nothing
End If

Function UrlPagina(nuevaPagina, nuevoLimite)
    Dim qs
    qs = "buscado=1"
    If EsCadenaNoVacia(fCod)    Then qs = qs & "&codSolicitud="  & Server.URLEncode(fCod)
    If EsCadenaNoVacia(fImei)   Then qs = qs & "&imei="          & Server.URLEncode(fImei)
    If EsCadenaNoVacia(fMsisdn) Then qs = qs & "&msisdn="        & Server.URLEncode(fMsisdn)
    If EsCadenaNoVacia(fTipo)   And fTipo <> "-1" Then qs = qs & "&tipo="    & Server.URLEncode(fTipo)
    If EsCadenaNoVacia(fEstado) And fEstado <> "-1" Then qs = qs & "&estado=" & Server.URLEncode(fEstado)
    If EsCadenaNoVacia(fIdent)  Then qs = qs & "&identCliente="  & Server.URLEncode(fIdent)
    If EsCadenaNoVacia(fNombre) Then qs = qs & "&nombreCliente=" & Server.URLEncode(fNombre)
    If EsCadenaNoVacia(fDesde)  Then qs = qs & "&fechaDesde="    & Server.URLEncode(fDesde)
    If EsCadenaNoVacia(fHasta)  Then qs = qs & "&fechaHasta="    & Server.URLEncode(fHasta)
    qs = qs & "&limite=" & nuevoLimite & "&pagina=" & nuevaPagina
    UrlPagina = "/modules/solicitudes/consulta.asp?" & qs
End Function
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Solicitudes</title>
<link rel="stylesheet" href="/assets/css/site.css">
<link rel="icon" href="/images/ToperBTR_favicon.ico" type="image/x-icon">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Solicitudes</h2>

    <%
    ' formCancelar vive fuera de la tabla como form oculto; los checkboxes lo
    ' referencian con form="formCancelar" (HTML5) para evitar forms anidados.
    %>
    <form id="formCancelar" method="post" action="/modules/solicitudes/cancelar_confirmar.asp" style="display:none">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(token) %>">
    </form>

    <nav class="modulo-nav">
        <span class="modulo-nav-item activo">Buscar</span>
        <% If puedeEscribir Then %>
        <a href="/modules/solicitudes/pre_alta.asp" class="modulo-nav-item">+ Nueva solicitud</a>
        <% End If %>
        <% If buscado And puedeEscribir Then %>
        <button type="submit" form="formCancelar" class="modulo-nav-item modulo-nav-cancelar">&#x2715; Cancelar seleccionadas</button>
        <% End If %>
    </nav>

    <form method="get" action="/modules/solicitudes/consulta.asp" class="formulario-busqueda">
        <input type="hidden" name="buscado" value="1">
        <div class="formulario-fila">
            <label>Codigo <input type="text" name="codSolicitud" value="<%= Server.HTMLEncode(fCod) %>"></label>
            <label>IMEI <input type="text" name="imei" value="<%= Server.HTMLEncode(fImei) %>"></label>
            <label>MSISDN <input type="text" name="msisdn" value="<%= Server.HTMLEncode(fMsisdn) %>"></label>
            <label>NIF/CIF cliente <input type="text" name="identCliente" value="<%= Server.HTMLEncode(fIdent) %>"></label>
            <label>Nombre cliente <input type="text" name="nombreCliente" value="<%= Server.HTMLEncode(fNombre) %>"></label>            
            <label>Tipo
                <select name="tipo">
                    <option value="-1">Todos</option>
                    <% Dim iTipo
                    For iTipo = 0 To nTipos - 1 %>
                    <option value="<%= arrTipoCod(iTipo) %>"<% If CStr(fTipo) = CStr(arrTipoCod(iTipo)) Then Response.Write " selected" %>><%= Server.HTMLEncode(arrTipoDesc(iTipo)) %></option>
                    <% Next %>
                </select>
            </label>
            <label>Estado
                <select name="estado">
                    <option value="-1">Todos</option>
                    <% Dim iEstado
                    For iEstado = 0 To nEstados - 1 %>
                    <option value="<%= arrEstadoCod(iEstado) %>"<% If CStr(fEstado) = CStr(arrEstadoCod(iEstado)) Then Response.Write " selected" %>><%= Server.HTMLEncode(arrEstadoDesc(iEstado)) %></option>
                    <% Next %>
                </select>
            </label>
            <label>Registros
                <select name="limite">
                    <option value="25"<% If limite = 25 Then Response.Write " selected" %>>25</option>
                    <option value="100"<% If limite = 100 Then Response.Write " selected" %>>100</option>
                    <option value="1000"<% If limite = 1000 Then Response.Write " selected" %>>1000</option>
                </select>
            </label>
        </div>
        <div class="formulario-fila">
            <label>Creacion desde <input type="date" name="fechaDesde" value="<%= Server.HTMLEncode(fDesde) %>"></label>
            <label>Creacion hasta <input type="date" name="fechaHasta" value="<%= Server.HTMLEncode(fHasta) %>"></label>
        </div>
        <div class="formulario-fila">
            <button type="submit">Filtrar</button>
        </div>
    </form>

    <% If buscado Then %>

    <div class="paginacion">
        <span class="paginacion-info">
            <%= totalRegistros %> resultado<% If totalRegistros <> 1 Then Response.Write "s" %>
            &mdash; P&aacute;gina <%= pagina %> de <%= totalPaginas %>
        </span>
        <nav class="paginacion-nav" aria-label="Paginaci&oacute;n">
            <% If pagina > 1 Then %>
            <a href="<%= UrlPagina(pagina - 1, limite) %>" class="paginacion-link">&larr; Anterior</a>
            <% Else %>
            <span class="paginacion-link deshabilitado">&larr; Anterior</span>
            <% End If %>
            <% If pagina < totalPaginas Then %>
            <a href="<%= UrlPagina(pagina + 1, limite) %>" class="paginacion-link">Siguiente &rarr;</a>
            <% Else %>
            <span class="paginacion-link deshabilitado">Siguiente &rarr;</span>
            <% End If %>
        </nav>
    </div>

    <table class="tabla-datos">
    <thead>
    <tr>
        <% If puedeEscribir Then %><th></th><% End If %>
        <th>Codigo</th><th>Tipo</th><th>IMEI</th><th>MSISDN</th><th>Cliente</th><th>Estado</th><th>Creacion</th>
    </tr>
    </thead>
    <tbody>
    <% If nFilas = 0 Then %>
    <tr><td colspan="<% If puedeEscribir Then Response.Write "8" Else Response.Write "7" End If %>" style="text-align:center;color:#777;padding:1.5rem 0">Sin resultados para los filtros indicados.</td></tr>
    <% Else %>
    <% Dim fi
    For fi = 0 To nFilas - 1 %>
    <tr>
        <% If puedeEscribir Then %>
        <td>
            <% If EsSolicitudCancelable(conn, aIndice(fi)) Then %>
            <input type="checkbox" name="seleccion" value="<%= aIndice(fi) %>" form="formCancelar">
            <% Else %>
            <input type="checkbox" disabled title="Esta solicitud no puede cancelarse">
            <% End If %>
        </td>
        <% End If %>
        <td><a href="/modules/solicitudes/detalle.asp?cod=<%= aIndice(fi) %>"><%= Server.HTMLEncode(aCodSol(fi)) %></a></td>
        <td><%= Server.HTMLEncode(aTipo(fi)) %></td>
        <td><%= Server.HTMLEncode(aImei(fi)) %></td>
        <td><%= Server.HTMLEncode(aMsisdn(fi)) %></td>
        <td><%= Server.HTMLEncode(aCliente(fi)) %></td>
        <td><%= Server.HTMLEncode(aEstado(fi)) %></td>
        <td><%= aFecha(fi) %></td>
    </tr>
    <% Next %>
    <% End If %>
    </tbody>
    </table>

    <% End If %>
</main>
</div>
<!--#include virtual="/include/layout/footer.asp"-->
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>
