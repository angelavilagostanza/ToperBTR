<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/incidencias_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequiereAccion(conn, CodigoAccion(conn, "LECTURA_INCIDENCIAS"))

Dim fCod, fTipo, fEstado, fOperador, fCodSolicitud, fNombreFichero, fDesde, fHasta
Dim buscado, limite, pagina, totalRegistros, totalPaginas, offsetReg

fCod           = Trim(Request.QueryString("codIncidencia"))
fTipo          = Request.QueryString("tipo")
fEstado        = Request.QueryString("estado")
fOperador      = Request.QueryString("operador")
fCodSolicitud  = Trim(Request.QueryString("codSolicitud"))
fNombreFichero = Trim(Request.QueryString("nombreFichero"))
fDesde         = Trim(Request.QueryString("fechaDesde"))
fHasta         = Trim(Request.QueryString("fechaHasta"))
buscado        = (Request.QueryString("buscado") = "1")

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

Dim rsIncidencias, rsTipos, rsEstados, rsOperadores
Set rsTipos      = ListarTiposIncidencia(conn)
Set rsEstados    = ListarEstadosIncidencia(conn)
Set rsOperadores = ListarOperadores(conn)

If buscado Then
    totalRegistros = ContarIncidencias(conn, fCod, fTipo, fEstado, fOperador, fCodSolicitud, fNombreFichero, fDesde, fHasta)
    totalPaginas   = Int((totalRegistros + limite - 1) / limite)
    If totalPaginas < 1 Then totalPaginas = 1
    If pagina > totalPaginas Then pagina = totalPaginas
    offsetReg      = (pagina - 1) * limite
    Set rsIncidencias = BuscarIncidencias(conn, fCod, fTipo, fEstado, fOperador, fCodSolicitud, fNombreFichero, fDesde, fHasta, limite, offsetReg)
End If

Function UrlPagina(nuevaPagina, nuevoLimite)
    Dim qs
    qs = "buscado=1"
    If EsCadenaNoVacia(fCod) Then qs = qs & "&codIncidencia=" & Server.URLEncode(fCod)
    If EsCadenaNoVacia(fTipo) And fTipo <> "-1" Then qs = qs & "&tipo=" & Server.URLEncode(fTipo)
    If EsCadenaNoVacia(fEstado) And fEstado <> "-1" Then qs = qs & "&estado=" & Server.URLEncode(fEstado)
    If EsCadenaNoVacia(fOperador) And fOperador <> "-1" Then qs = qs & "&operador=" & Server.URLEncode(fOperador)
    If EsCadenaNoVacia(fCodSolicitud) Then qs = qs & "&codSolicitud=" & Server.URLEncode(fCodSolicitud)
    If EsCadenaNoVacia(fNombreFichero) Then qs = qs & "&nombreFichero=" & Server.URLEncode(fNombreFichero)
    If EsCadenaNoVacia(fDesde) Then qs = qs & "&fechaDesde=" & Server.URLEncode(fDesde)
    If EsCadenaNoVacia(fHasta) Then qs = qs & "&fechaHasta=" & Server.URLEncode(fHasta)
    qs = qs & "&limite=" & nuevoLimite & "&pagina=" & nuevaPagina
    UrlPagina = "/modules/incidencias/buscar.asp?" & qs
End Function
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Incidencias</title>
<link rel="stylesheet" href="/assets/css/site.css">
<link rel="icon" href="/images/ToperBTR_favicon.ico" type="image/x-icon">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Incidencias</h2>

    <form method="get" action="/modules/incidencias/buscar.asp" class="formulario-busqueda" onsubmit="return validarFiltros(this)">
        <input type="hidden" name="buscado" value="1">
        <div class="formulario-fila">
            <label>Codigo <input type="text" name="codIncidencia" value="<%= Server.HTMLEncode(fCod) %>"></label>
            <label>Tipo
                <select name="tipo">
                    <option value="-1">Todos</option>
                    <% Do While Not rsTipos.EOF %>
                    <option value="<%= rsTipos("COD_TIPO_INCIDENCIA") %>"<% If CStr(fTipo) = CStr(rsTipos("COD_TIPO_INCIDENCIA")) Then Response.Write " selected" %>><%= Server.HTMLEncode(rsTipos("DESCRIPCION")) %></option>
                    <% rsTipos.MoveNext
                    Loop %>
                </select>
            </label>
            <label>Estado
                <select name="estado">
                    <option value="-1">Todos</option>
                    <% Do While Not rsEstados.EOF %>
                    <option value="<%= rsEstados("COD_ESTADO") %>"<% If CStr(fEstado) = CStr(rsEstados("COD_ESTADO")) Then Response.Write " selected" %>><%= Server.HTMLEncode(rsEstados("DESCRIPCION")) %></option>
                    <% rsEstados.MoveNext
                    Loop %>
                </select>
            </label>
            <label>Operador
                <select name="operador">
                    <option value="-1">Todos</option>
                    <% Do While Not rsOperadores.EOF %>
                    <option value="<%= rsOperadores("COD_OPERADOR") %>"<% If CStr(fOperador) = CStr(rsOperadores("COD_OPERADOR")) Then Response.Write " selected" %>><%= Server.HTMLEncode(rsOperadores("NOMBRE")) %></option>
                    <% rsOperadores.MoveNext
                    Loop %>
                </select>
            </label>
            <label>Codigo solicitud <input type="text" name="codSolicitud" value="<%= Server.HTMLEncode(fCodSolicitud) %>"></label>
            <label>Nombre fichero <input type="text" name="nombreFichero" value="<%= Server.HTMLEncode(fNombreFichero) %>"></label>
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
        <div id="aviso-filtro" class="flash-info" hidden style="margin-top:.5rem">
            Indica al menos un criterio de b&uacute;squeda. Puedes usar <strong>Creaci&oacute;n desde</strong> para listar incidencias por fecha.
        </div>
        <div class="formulario-fila">
            <button type="submit">Filtrar</button>
        </div>
    </form>
    <script>
    function validarFiltros(f) {
        var textos = ['codIncidencia','codSolicitud','nombreFichero','fechaDesde','fechaHasta'];
        for (var i = 0; i < textos.length; i++) {
            if (f[textos[i]].value.trim() !== '') { document.getElementById('aviso-filtro').hidden = true; return true; }
        }
        document.getElementById('aviso-filtro').hidden = false;
        return false;
    }
    </script>

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
    <tr><th>Codigo</th><th>Tipo</th><th>Operador</th><th>Estado</th><th>Vinculada a</th><th>Creacion</th></tr>
    </thead>
    <tbody>
    <% If rsIncidencias.EOF Then %>
    <tr><td colspan="6" style="text-align:center;color:#777;padding:1.5rem 0">Sin resultados para los filtros indicados.</td></tr>
    <% Else %>
    <% Do While Not rsIncidencias.EOF %>
    <tr>
        <td><a href="/modules/incidencias/detalle.asp?cod=<%= rsIncidencias("COD_IND_INCIDENCIA") %>"><%= Server.HTMLEncode(rsIncidencias("COD_INCIDENCIA")) %></a></td>
        <td><%= Server.HTMLEncode(rsIncidencias("TIPO_INCIDENCIA")) %></td>
        <td><%= Server.HTMLEncode(rsIncidencias("OPERADOR")) %></td>
        <td><%= Server.HTMLEncode(rsIncidencias("ESTADO")) %></td>
        <td>
            <% If rsIncidencias("TIPO_OBJETO") = "S" Then %>
            Solicitud <%= Server.HTMLEncode(rsIncidencias("COD_SOLICITUD") & "") %>
            <% ElseIf rsIncidencias("TIPO_OBJETO") = "F" Then %>
            Fichero <%= Server.HTMLEncode(rsIncidencias("NOMBRE_FICHERO") & "") %>
            <% Else %>
            &mdash;
            <% End If %>
        </td>
        <td><%= rsIncidencias("FECHA_CREACION") %></td>
    </tr>
    <% rsIncidencias.MoveNext
    Loop %>
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
If buscado Then rsIncidencias.Close
rsTipos.Close
rsEstados.Close
rsOperadores.Close
conn.Close
Set conn = Nothing
%>
