<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/auditoria_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequierePerfil(Array(CodigoPerfil(conn, "Seguridad")))

Dim fUsuario, fAccion, fDesde, fHasta, fDetalle
Dim buscado, limite, pagina, totalRegistros, totalPaginas, offsetReg

fUsuario = Trim(Request.QueryString("usuario"))
fAccion  = Request.QueryString("accion")
fDesde   = Trim(Request.QueryString("fechaDesde"))
fHasta   = Trim(Request.QueryString("fechaHasta"))
fDetalle = Trim(Request.QueryString("detalle"))
buscado  = (Request.QueryString("buscado") = "1")

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

Dim rsAuditoria, rsAcciones
Set rsAcciones = ListarAccionesAuditoria(conn)

If buscado Then
    totalRegistros = ContarAuditoria(conn, fUsuario, fAccion, fDesde, fHasta, fDetalle)
    totalPaginas   = Int((totalRegistros + limite - 1) / limite)
    If totalPaginas < 1 Then totalPaginas = 1
    If pagina > totalPaginas Then pagina = totalPaginas
    offsetReg      = (pagina - 1) * limite
    Set rsAuditoria = BuscarAuditoria(conn, fUsuario, fAccion, fDesde, fHasta, fDetalle, limite, offsetReg)
End If

Function UrlPagina(nuevaPagina, nuevoLimite)
    Dim qs
    qs = "buscado=1"
    If EsCadenaNoVacia(fUsuario) Then qs = qs & "&usuario=" & Server.URLEncode(fUsuario)
    If EsCadenaNoVacia(fAccion)  Then qs = qs & "&accion="  & Server.URLEncode(fAccion)
    If EsCadenaNoVacia(fDesde)   Then qs = qs & "&fechaDesde=" & Server.URLEncode(fDesde)
    If EsCadenaNoVacia(fHasta)   Then qs = qs & "&fechaHasta=" & Server.URLEncode(fHasta)
    If EsCadenaNoVacia(fDetalle) Then qs = qs & "&detalle=" & Server.URLEncode(fDetalle)
    qs = qs & "&limite=" & nuevoLimite & "&pagina=" & nuevaPagina
    UrlPagina = "/modules/auditoria/buscar.asp?" & qs
End Function
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Auditoria</title>
<link rel="stylesheet" href="/assets/css/site.css">
<link rel="icon" href="/images/ToperBTR_favicon.ico" type="image/x-icon">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Auditor&iacute;a</h2>

    <form method="get" action="/modules/auditoria/buscar.asp" class="formulario-busqueda" onsubmit="return validarFiltros(this)">
        <input type="hidden" name="buscado" value="1">
        <div class="formulario-fila">
            <label>Usuario <input type="text" maxlength="50" name="usuario" value="<%= Server.HTMLEncode(fUsuario) %>"></label>
            <label>Accion
                <select name="accion">
                    <option value="">Todas</option>
                    <% Do While Not rsAcciones.EOF %>
                    <option value="<%= Server.HTMLEncode(rsAcciones("ACCION")) %>"<% If fAccion = rsAcciones("ACCION") Then Response.Write " selected" %>><%= Server.HTMLEncode(rsAcciones("ACCION")) %></option>
                    <% rsAcciones.MoveNext
                    Loop %>
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
            <label>Fecha desde <input type="date" name="fechaDesde" value="<%= Server.HTMLEncode(fDesde) %>"></label>
            <label>Fecha hasta <input type="date" name="fechaHasta" value="<%= Server.HTMLEncode(fHasta) %>"></label>
            <label>Detalle contiene <input type="text" maxlength="100" name="detalle" value="<%= Server.HTMLEncode(fDetalle) %>" placeholder="solo junto a otro filtro"></label>
        </div>
        <div id="aviso-filtro" class="flash-info" hidden style="margin-top:.5rem">
            Indica al menos un criterio de b&uacute;squeda (Usuario, Accion o fecha). Puedes usar <strong>Fecha desde</strong> para listar la auditor&iacute;a por fecha. El campo <strong>Detalle</strong> solo se aplica junto a otro criterio.
        </div>
        <div class="formulario-fila">
            <button type="submit">Filtrar</button>
        </div>
    </form>
    <script>
    function validarFiltros(f) {
        var textos = ['usuario','accion','fechaDesde','fechaHasta'];
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
    <tr><th>Fecha</th><th>Usuario</th><th>Accion</th><th>Entidad</th><th>Resultado</th><th>Detalle</th><th>IP</th></tr>
    </thead>
    <tbody>
    <% If rsAuditoria.EOF Then %>
    <tr><td colspan="7" style="text-align:center;color:#777;padding:1.5rem 0">Sin resultados para los filtros indicados.</td></tr>
    <% Else %>
    <% Do While Not rsAuditoria.EOF %>
    <tr>
        <td><%= rsAuditoria("FECHA") %></td>
        <td><%= SiVerdadero(EsCadenaNoVacia(rsAuditoria("COD_USUARIO") & ""), Server.HTMLEncode(rsAuditoria("COD_USUARIO") & ""), "(sistema)") %></td>
        <td><code><%= Server.HTMLEncode(rsAuditoria("ACCION") & "") %></code></td>
        <td><%= Server.HTMLEncode(rsAuditoria("ENTIDAD") & "") %><% If EsCadenaNoVacia(rsAuditoria("ENTIDAD_ID") & "") Then %> #<%= Server.HTMLEncode(rsAuditoria("ENTIDAD_ID") & "") %><% End If %></td>
        <td><%= Server.HTMLEncode(rsAuditoria("RESULTADO") & "") %></td>
        <td><%= Server.HTMLEncode(rsAuditoria("DETALLE") & "") %></td>
        <td><%= Server.HTMLEncode(rsAuditoria("IP") & "") %></td>
    </tr>
    <% rsAuditoria.MoveNext
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
If buscado Then rsAuditoria.Close
rsAcciones.Close
conn.Close
Set conn = Nothing
%>
