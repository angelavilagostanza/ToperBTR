<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/solicitudes_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequiereAccion(conn, CodigoAccion(conn, "ESCRITURA_SOLICITUDES"))

If Not TokenCsrfValido(Request.Form("csrf")) Then
    conn.Close
    Response.Redirect "/errors/error.asp?msg=csrf"
    Response.End
End If

Dim seleccionCsv
seleccionCsv = Request.Form("seleccion")

If Not EsCadenaNoVacia(seleccionCsv) Then
    conn.Close
    Set conn = Nothing
    Call EstablecerMensajeFlash("error", "No se ha seleccionado ninguna solicitud.")
    Response.Redirect "/modules/solicitudes/consulta.asp"
    Response.End
End If

Dim listaIds, i
listaIds = Split(seleccionCsv, ",")

Dim tokenConfirm
tokenConfirm = GenerarTokenCsrf()
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Confirmar cancelacion</title>
<link rel="stylesheet" href="/assets/css/site.css">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Confirmar cancelacion de solicitud(es)</h2>
    <ul>
    <% For i = 0 To UBound(listaIds)
        Dim s
        Set s = ObtenerSolicitudPorIndice(conn, CLng(listaIds(i)))
        If Not s Is Nothing Then %>
        <li><%= Server.HTMLEncode(s("CodSolicitud")) %> &mdash; IMEI <%= Server.HTMLEncode(s("Imei")) %></li>
        <% End If
    Next %>
    </ul>
    <form method="post" action="/modules/solicitudes/cancelar_ejecutar.asp">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(tokenConfirm) %>">
        <input type="hidden" name="seleccion" value="<%= Server.HTMLEncode(seleccionCsv) %>">
        <button type="submit">Confirmar cancelacion</button>
        <a href="/modules/solicitudes/consulta.asp">Cancelar</a>
    </form>
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
