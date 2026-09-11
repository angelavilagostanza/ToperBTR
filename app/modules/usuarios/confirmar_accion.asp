<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/usuarios_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequierePerfil(Array(CodigoPerfil(conn, "Seguridad")))

If Not TokenCsrfValido(Request.Form("csrf")) Then
    conn.Close
    Response.Redirect "/errors/error.asp?msg=csrf"
    Response.End
End If

Dim accion, seleccionCsv
accion = Request.Form("accion")
seleccionCsv = Request.Form("seleccion")

If Not EsCadenaNoVacia(seleccionCsv) Then
    conn.Close
    Call EstablecerMensajeFlash("error", "No se ha seleccionado ningun usuario.")
    Response.Redirect "/modules/usuarios/listar.asp"
    Response.End
End If

Dim listaIds, i, textoAccion
listaIds = Split(seleccionCsv, ",")
If accion = "baja" Then
    textoAccion = "dar de baja"
Else
    textoAccion = "desbloquear"
End If

Dim tokenConfirm
tokenConfirm = GenerarTokenCsrf()
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Confirmar</title>
<link rel="stylesheet" href="/assets/css/site.css">
<link rel="icon" href="/images/ToperBTR_favicon.ico" type="image/x-icon">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Confirmar: <%= Server.HTMLEncode(textoAccion) %></h2>
    <p>Se va a aplicar esta accion a los siguientes usuarios:</p>
    <ul>
    <% For i = 0 To UBound(listaIds)
        Dim u
        Set u = ObtenerUsuarioPorIndice(conn, CLng(listaIds(i)))
        If Not u Is Nothing Then %>
        <li><%= Server.HTMLEncode(u("CodUsuario")) %> &mdash; <%= Server.HTMLEncode(u("Nombre")) %> <%= Server.HTMLEncode(u("Apellidos")) %></li>
        <% End If
    Next %>
    </ul>
    <form method="post" action="/modules/usuarios/accion_ejecutar.asp">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(tokenConfirm) %>">
        <input type="hidden" name="accion" value="<%= Server.HTMLEncode(accion) %>">
        <input type="hidden" name="seleccion" value="<%= Server.HTMLEncode(seleccionCsv) %>">
        <button type="submit">Confirmar</button>
        <a href="/modules/usuarios/listar.asp">Cancelar</a>
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
