<!--#include virtual="/include/bootstrap.asp"-->
<%
Dim conn
Set conn = NuevaConexion()

Dim codAdmin, codTramitacion
codAdmin = CodigoPerfil(conn, "Administrador")
codTramitacion = CodigoPerfil(conn, "Tramitación")
Call RequierePerfil(Array(codAdmin, codTramitacion))

Dim rsRazones, tokenCsrf
Set rsRazones = ListarRazones(conn)
tokenCsrf = GenerarTokenCsrf()
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Bloquear IMEI</title>
<link rel="stylesheet" href="/assets/css/site.css">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Bloquear IMEI</h2>

    <form method="post" action="/modules/bloqueos/bloquear_do.asp" class="form-edicion">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(tokenCsrf) %>">

        <div class="campo">
            <label for="imei">IMEI <span class="obligatorio">*</span></label>
            <input type="text" id="imei" name="imei" maxlength="15" required
                placeholder="15 digitos numericos" autocomplete="off">
            <span class="ayuda-campo">Debe ser un IMEI de un cliente activo de Grupo MASMOVIL.</span>
        </div>

        <div class="campo">
            <label for="razon">Razon de bloqueo <span class="obligatorio">*</span></label>
            <select id="razon" name="razon" required>
                <option value="">-- Seleccione una razon --</option>
                <% Do While Not rsRazones.EOF %>
                <option value="<%= CLng(rsRazones("COD_RAZON")) %>"><%= Server.HTMLEncode(rsRazones("DESCRIPCION") & "") %></option>
                <% rsRazones.MoveNext : Loop %>
            </select>
        </div>

        <div class="campo">
            <label for="comentario">Comentario</label>
            <textarea id="comentario" name="comentario" maxlength="500" rows="3"
                placeholder="Informacion adicional (opcional)"></textarea>
        </div>

        <div class="acciones-form">
            <button type="submit" class="btn btn-peligro"
                onclick="return confirm('&iquest;Confirma el bloqueo del IMEI ' + document.getElementById(\'imei\').value + '?')">
                Bloquear IMEI
            </button>
            <a href="/modules/bloqueos/bloqueo.asp" class="btn btn-secundario">Cancelar</a>
        </div>
    </form>

</main>
</div>
</div>
</body>
</html>
<%
rsRazones.Close
conn.Close
Set conn = Nothing
%>
