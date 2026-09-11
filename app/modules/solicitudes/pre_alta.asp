<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/solicitudes_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequiereAccion(conn, CodigoAccion(conn, "ESCRITURA_SOLICITUDES"))

Dim token
token = GenerarTokenCsrf()

Dim errorMsg
errorMsg = Request.QueryString("error")

Dim rsTipos
Set rsTipos = ListarTiposSolicitud(conn)
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Nueva solicitud</title>
<link rel="stylesheet" href="/assets/css/site.css">
<link rel="icon" href="/images/ToperBTR_favicon.ico" type="image/x-icon">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Nueva solicitud</h2>
    <% If EsCadenaNoVacia(errorMsg) Then %>
    <p class="flash-error"><%= Server.HTMLEncode(errorMsg) %></p>
    <% End If %>
    <form method="post" action="/modules/solicitudes/pre_alta_do.asp" class="formulario-vertical">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(token) %>">
        <label>IMEI
            <input type="text" name="imei" maxlength="15" required>
        </label>
        <label>Tipo de solicitud
            <select name="tipo" required>
                <option value="">-- Seleccione --</option>
                <% Do While Not rsTipos.EOF %>
                <option value="<%= rsTipos("CLAVE") %>"><%= Server.HTMLEncode(rsTipos("DESCRIPCION")) %></option>
                <% rsTipos.MoveNext
                Loop %>
            </select>
        </label>
        <button type="submit">Continuar</button>
    </form>
</main>
</div>
<!--#include virtual="/include/layout/footer.asp"-->
</div>
</body>
</html>
<%
rsTipos.Close
conn.Close
Set conn = Nothing
%>
