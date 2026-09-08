<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/bloqueos_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()

' Gate explícito por perfil — el legacy controlaba esto solo ocultando el enlace de
' menú (sin comprobación real en servidor). Ver docs/04-matriz-permisos.md.
Call RequierePerfil(Array(CodigoPerfilCacheado("Administrador"), CodigoPerfilCacheado("Tramitación")))

Dim fImei, fTipo, fRazon, fDesde, fHasta, busquedaRealizada
fImei    = Trim(Request.Form("imei") & "")
fTipo    = Trim(Request.Form("tipo") & "")
fRazon   = Trim(Request.Form("razon") & "")
fDesde   = Trim(Request.Form("fechaDesde") & "")
fHasta   = Trim(Request.Form("fechaHasta") & "")
busquedaRealizada = (Request.ServerVariables("REQUEST_METHOD") = "POST")

' Validar IMEI si se ha introducido: 15 digitos numericos exactos.
' Corrige el bug real de BusquedaBloqueoForm.validate() del legacy, donde la comprobacion
' de que el IMEI fuera numerico estaba en un else if con condicion duplicada — codigo muerto
' que nunca se ejecutaba, dejando sin validar el formato numerico.
Dim errImei
errImei = ""
If busquedaRealizada And fImei <> "" Then
    If Len(fImei) <> 15 Then
        errImei = "El IMEI debe tener exactamente 15 digitos."
    ElseIf Not CumpleFormato(fImei, "^\d{15}$") Then
        errImei = "El IMEI debe contener solo digitos numericos."
    End If
End If

' Validar fechas si se han introducido
Dim errFecha
errFecha = ""
If busquedaRealizada Then
    If fDesde <> "" And Not IsDate(fDesde) Then
        errFecha = "La fecha 'desde' no tiene un formato valido (dd/mm/aaaa)."
    ElseIf fHasta <> "" And Not IsDate(fHasta) Then
        errFecha = "La fecha 'hasta' no tiene un formato valido (dd/mm/aaaa)."
    ElseIf fDesde <> "" And fHasta <> "" And IsDate(fDesde) And IsDate(fHasta) Then
        If CDate(fDesde) > CDate(fHasta) Then
            errFecha = "La fecha 'desde' no puede ser posterior a la fecha 'hasta'."
        End If
    End If
End If

Dim hayErrores
hayErrores = (errImei <> "" Or errFecha <> "")

Dim rsBloqueos, rsTipos, rsRazones
Set rsTipos  = ListarTiposBloqueo(conn)
Set rsRazones = ListarRazones(conn)

If busquedaRealizada And Not hayErrores Then
    Set rsBloqueos = ListarBloqueos(conn, fImei, fTipo, fRazon, _
        SiVerdadero(fDesde <> "" And IsDate(fDesde), fDesde, ""), _
        SiVerdadero(fHasta <> "" And IsDate(fHasta), fHasta, ""))
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
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Bloqueos</title>
<link rel="stylesheet" href="/assets/css/site.css">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Bloqueos</h2>

    <div class="acciones-cabecera">
        <a href="/modules/bloqueos/bloquear.asp" class="btn btn-primario">Bloquear IMEI</a>
    </div>

    <form method="post" action="/modules/bloqueos/bloqueo.asp" class="form-busqueda">
        <fieldset>
            <legend>Filtros de busqueda</legend>
            <div class="campo">
                <label for="imei">IMEI</label>
                <input type="text" id="imei" name="imei" value="<%= Server.HTMLEncode(fImei) %>" maxlength="15" placeholder="15 digitos">
                <% If errImei <> "" Then %><span class="error-campo"><%= Server.HTMLEncode(errImei) %></span><% End If %>
            </div>
            <div class="campo">
                <label for="tipo">Tipo de bloqueo</label>
                <select id="tipo" name="tipo">
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
            </div>
            <div class="campo">
                <label for="razon">Razon</label>
                <select id="razon" name="razon">
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
            </div>
            <div class="campo">
                <label for="fechaDesde">Fecha inclusion desde</label>
                <input type="text" id="fechaDesde" name="fechaDesde" value="<%= Server.HTMLEncode(fDesde) %>" maxlength="10" placeholder="dd/mm/aaaa">
            </div>
            <div class="campo">
                <label for="fechaHasta">Fecha inclusion hasta</label>
                <input type="text" id="fechaHasta" name="fechaHasta" value="<%= Server.HTMLEncode(fHasta) %>" maxlength="10" placeholder="dd/mm/aaaa">
                <% If errFecha <> "" Then %><span class="error-campo"><%= Server.HTMLEncode(errFecha) %></span><% End If %>
            </div>
        </fieldset>
        <div class="acciones-form">
            <button type="submit" class="btn btn-primario">Buscar</button>
        </div>
    </form>

    <% If busquedaRealizada And Not hayErrores Then %>
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
