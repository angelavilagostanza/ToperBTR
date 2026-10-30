<%@ CODEPAGE=65001 LANGUAGE="VBScript" %>
<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/usuarios_dal.asp"-->
<%
Response.CodePage = 65001
Response.CharSet = "utf-8"
Dim conn
Set conn = NuevaConexion()
Call RequierePerfil(Array(CodigoPerfil(conn, "Seguridad")))

If Not TokenCsrfValido(Request.Form("csrf")) Then
    conn.Close
    Response.Redirect "/errors/error.asp?msg=csrf"
    Response.End
End If

Dim accion, seleccionCsv, volver
accion = Request.Form("accion")
seleccionCsv = Request.Form("seleccion")

' "volver" trae los filtros/pagina/limite actuales de listar.asp para no perderlos al
' terminar la accion. Viene de un campo de formulario (manipulable), asi que se valida
' que sea una ruta local a listar.asp antes de usarla en ningun sitio - evita un
' redirect abierto a una URL externa.
volver = Request.Form("volver")
If Not (EsCadenaNoVacia(volver) And Left(volver, Len("/modules/usuarios/listar.asp")) = "/modules/usuarios/listar.asp") Then
    volver = "/modules/usuarios/listar.asp"
End If

If Not EsCadenaNoVacia(seleccionCsv) Then
    conn.Close
    Call EstablecerMensajeFlash("error", "No se ha seleccionado ningún usuario.")
    Response.Redirect volver
    Response.End
End If

Dim listaIds, i, textoAccion
listaIds = Split(seleccionCsv, ",")

Select Case accion
    Case "baja"
        textoAccion = "dar de baja"
    Case "bloqueo"
        textoAccion = "bloquear"
    Case "desbloqueo"
        textoAccion = "desbloquear"
    Case Else
        conn.Close
        Response.Redirect volver
        Response.End
End Select

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
    <p>Se va a aplicar esta acción a los siguientes usuarios:</p>
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
        <input type="hidden" name="volver" value="<%= Server.HTMLEncode(volver) %>">
        <button type="submit">Confirmar</button>
        <a href="<%= Server.HTMLEncode(volver) %>">Cancelar</a>
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
