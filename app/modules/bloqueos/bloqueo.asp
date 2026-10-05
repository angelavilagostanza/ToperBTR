<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/bloqueos_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()

' Gate explícito por perfil — el legacy controlaba esto solo ocultando el enlace de
' menú (sin comprobación real en servidor). Ver docs/04-matriz-permisos.md.
Call RequierePerfil(Array(CodigoPerfilCacheado("Administrador"), CodigoPerfilCacheado("Tramitación")))

Dim fImei, fTipo, fRazon, fDesde, fHasta, buscado
Dim limite, pagina, totalRegistros, totalPaginas, offsetReg

fImei    = Trim(Request.QueryString("imei") & "")
fTipo    = Trim(Request.QueryString("tipo") & "")
fRazon   = Trim(Request.QueryString("razon") & "")
fDesde   = Trim(Request.QueryString("fechaDesde") & "")
fHasta   = Trim(Request.QueryString("fechaHasta") & "")
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

' Validar IMEI si se ha introducido: 15 digitos numericos exactos.
' Corrige el bug real de BusquedaBloqueoForm.validate() del legacy, donde la comprobacion
' de que el IMEI fuera numerico estaba en un else if con condicion duplicada — codigo muerto
' que nunca se ejecutaba, dejando sin validar el formato numerico.
Dim errImei
errImei = ""
If buscado And fImei <> "" Then
    If Len(fImei) <> 15 Then
        errImei = "El IMEI debe tener exactamente 15 digitos."
    ElseIf Not CumpleFormato(fImei, "^\d{15}$") Then
        errImei = "El IMEI debe contener solo digitos numericos."
    End If
End If

Dim hayErrores
hayErrores = (errImei <> "")

Dim rsBloqueos, rsTipos, rsRazones
Set rsTipos  = ListarTiposBloqueo(conn)
Set rsRazones = ListarRazones(conn)

' Los campos de fecha son <input type="date">: el navegador ya garantiza formato
' ISO "aaaa-mm-dd" o cadena vacia. ContarBloqueos/ListarBloqueos filtran con
' IsDate()/CDate(), igual que el resto de busquedas de la aplicacion
' (incidencias_dal.asp, ficheros_dal.asp, solicitudes_dal.asp).
If buscado And Not hayErrores Then
    totalRegistros = ContarBloqueos(conn, fImei, fTipo, fRazon, fDesde, fHasta)
    totalPaginas   = Int((totalRegistros + limite - 1) / limite)
    If totalPaginas < 1 Then totalPaginas = 1
    If pagina > totalPaginas Then pagina = totalPaginas
    offsetReg      = (pagina - 1) * limite

    Set rsBloqueos = ListarBloqueos(conn, fImei, fTipo, fRazon, fDesde, fHasta, limite, offsetReg)
End If

' Arrays para los desplegables (el recordset es forward-only, no admite MoveFirst)
Dim mapaTipos, mapaRazones
Set mapaTipos  = Server.CreateObject("Scripting.Dictionary")
Set mapaRazones = Server.CreateObject("Scripting.Dictionary")
' Rellenar rsTipos en el mapa para poder pintar el desplegable y etiquetar filas
Do While Not rsTipos.EOF
    mapaTipos.Add CLng(rsTipos("COD_TIPO_BLOQUEO")) & "", rsTipos("DESCRIPCION") & ""
    rsTipos.MoveNext
Loop
rsTipos.Close
Do While Not rsRazones.EOF
    mapaRazones.Add CLng(rsRazones("COD_RAZON")) & "", rsRazones("DESCRIPCION") & ""
    rsRazones.MoveNext
Loop
rsRazones.Close

' Token CSRF para el formulario de desbloqueo
Dim tokenDesbloqueo
tokenDesbloqueo = GenerarTokenCsrf()

Function UrlPagina(nuevaPagina, nuevoLimite)
    Dim qs
    qs = "buscado=1"
    If EsCadenaNoVacia(fImei)  Then qs = qs & "&imei="       & Server.URLEncode(fImei)
    If EsCadenaNoVacia(fTipo)  Then qs = qs & "&tipo="       & Server.URLEncode(fTipo)
    If EsCadenaNoVacia(fRazon) Then qs = qs & "&razon="      & Server.URLEncode(fRazon)
    If EsCadenaNoVacia(fDesde) Then qs = qs & "&fechaDesde=" & Server.URLEncode(fDesde)
    If EsCadenaNoVacia(fHasta) Then qs = qs & "&fechaHasta=" & Server.URLEncode(fHasta)
    qs = qs & "&limite=" & nuevoLimite & "&pagina=" & nuevaPagina
    UrlPagina = "/modules/bloqueos/bloqueo.asp?" & qs
End Function
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Bloqueos</title>
<link rel="stylesheet" href="/assets/css/site.css">
<link rel="icon" href="/images/ToperBTR_favicon.ico" type="image/x-icon">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Bloqueos</h2>

    <nav class="modulo-nav">
        <span class="modulo-nav-item activo">Buscar</span>
        <a href="/modules/bloqueos/bloquear.asp" class="modulo-nav-item">+ Nuevo Bloqueo</a>
    </nav>

    <form method="get" action="/modules/bloqueos/bloqueo.asp" class="formulario-busqueda" onsubmit="return validarFiltros(this)">
        <input type="hidden" name="buscado" value="1">
        <div class="formulario-fila">
            <label>IMEI
                <input type="text" name="imei" value="<%= Server.HTMLEncode(fImei) %>" maxlength="15" placeholder="15 digitos">
                <% If errImei <> "" Then %><span class="error-campo"><%= Server.HTMLEncode(errImei) %></span><% End If %>
            </label>
            <label>Tipo de bloqueo
                <select name="tipo">
                    <option value="">-- Todos --</option>
                    <%
                    Dim clTipo
                    For Each clTipo In mapaTipos.Keys
                    %>
                    <option value="<%= Server.HTMLEncode(clTipo) %>"<% If fTipo = clTipo Then %> selected<% End If %>><%= Server.HTMLEncode(mapaTipos(clTipo)) %></option>
                    <%
                    Next
                    %>
                </select>
            </label>
            <label>Razon
                <select name="razon">
                    <option value="">-- Todas --</option>
                    <%
                    Dim clRazon
                    For Each clRazon In mapaRazones.Keys
                    %>
                    <option value="<%= Server.HTMLEncode(clRazon) %>"<% If fRazon = clRazon Then %> selected<% End If %>><%= Server.HTMLEncode(mapaRazones(clRazon)) %></option>
                    <%
                    Next
                    %>
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
            <label>Fecha inclusion desde
                <input type="date" name="fechaDesde" value="<%= Server.HTMLEncode(fDesde) %>">
            </label>
            <label>Fecha inclusion hasta
                <input type="date" name="fechaHasta" value="<%= Server.HTMLEncode(fHasta) %>">
            </label>
        </div>
        <div id="aviso-filtro" class="flash-info" hidden style="margin-top:.5rem">
            Indica al menos un criterio de b&uacute;squeda. Puedes usar <strong>Fecha inclusion desde</strong> para listar bloqueos por fecha.
        </div>
        <div class="formulario-fila">
            <button type="submit" class="btn btn-primario">Buscar</button>
        </div>
    </form>
    <script>
    function validarFiltros(f) {
        var textos = ['imei','fechaDesde','fechaHasta'];
        for (var i = 0; i < textos.length; i++) {
            if (f[textos[i]].value.trim() !== '') { document.getElementById('aviso-filtro').hidden = true; return true; }
        }
        document.getElementById('aviso-filtro').hidden = false;
        return false;
    }
    </script>

    <% If buscado And Not hayErrores Then %>

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

    <%
    If rsBloqueos.EOF Then
    %>
    <p class="aviso-vacio">No se han encontrado bloqueos con los filtros indicados.</p>
    <%
    Else
    %>
    <form method="post" action="/modules/bloqueos/desbloquear_do.asp">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(tokenDesbloqueo) %>">
        <table class="tabla-datos">
            <thead>
                <tr>
                    <th><input type="checkbox" id="selTodos" title="Seleccionar todos"></th>
                    <th>IMEI</th>
                    <th>Tipo de bloqueo</th>
                    <th>Razon</th>
                    <th>Fecha inclusion</th>
                    <th>Comentario</th>
                </tr>
            </thead>
            <tbody>
            <%
            Do While Not rsBloqueos.EOF
                Dim fila_imei, fila_tipo, fila_razon, fila_fecha, fila_comentario, fila_eneir
                fila_imei       = rsBloqueos("IMEI") & ""
                fila_tipo       = rsBloqueos("TIPO_BLOQUEO_DESC") & ""
                fila_razon      = rsBloqueos("RAZON_DESC") & ""
                fila_fecha      = rsBloqueos("FECHA_INCLUSION")
                fila_comentario = rsBloqueos("COMENTARIO") & ""
                fila_eneir      = CLng(rsBloqueos("EN_EIR") & "0")
            %>
                <tr>
                    <td><input type="checkbox" name="seleccion" value="<%= Server.HTMLEncode(fila_imei) %>"></td>
                    <td><code><%= Server.HTMLEncode(fila_imei) %></code></td>
                    <td><%= Server.HTMLEncode(fila_tipo) %></td>
                    <td><%= Server.HTMLEncode(fila_razon) %></td>
                    <td><%= SiVerdadero(IsDate(fila_fecha), FormatDateTime(fila_fecha, 2), "") %></td>
                    <td><%= Server.HTMLEncode(fila_comentario) %></td>
                </tr>
            <%
                rsBloqueos.MoveNext
            Loop
            rsBloqueos.Close
            %>
            </tbody>
        </table>
        <div class="acciones-form">
            <button type="submit" class="btn btn-peligro"
                onclick="return confirm('&iquest;Desbloquear los IMEI seleccionados? Esta accion es irreversible.')">
                Desbloquear seleccionados
            </button>
        </div>
    </form>
    <%
    End If
    %>
    <% End If %>

</main>
</div>
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>
<script>
document.getElementById('selTodos').addEventListener('change', function() {
    document.querySelectorAll('input[name="seleccion"]').forEach(function(cb) {
        cb.checked = document.getElementById('selTodos').checked;
    });
});
</script>
