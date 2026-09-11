<!--#include virtual="/include/bootstrap.asp"-->
<%
Call RequiereSesion()

If Not EsCadenaNoVacia(Session("Sol_Imei")) Then
    Response.Redirect "/modules/solicitudes/pre_alta.asp"
    Response.End
End If

Dim token
token = GenerarTokenCsrf()
Dim errorMsg
errorMsg = Request.QueryString("error")

Dim esInclusion, soloLecturaCliente, tituloTipo, atributoReadonly
esInclusion = (Session("Sol_TipoClave") = "I")
soloLecturaCliente = (Session("Sol_PrefillBloqueado") = True)
tituloTipo = SiVerdadero(esInclusion, "Inclusion", "Exclusion")
atributoReadonly = SiVerdadero(soloLecturaCliente, " readonly", "")
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
    <h2>Nueva solicitud &mdash; <%= Server.HTMLEncode(tituloTipo) %></h2>
    <p>IMEI: <strong><%= Server.HTMLEncode(Session("Sol_Imei")) %></strong></p>
    <% If EsCadenaNoVacia(errorMsg) Then %>
    <p class="flash-error"><%= Server.HTMLEncode(errorMsg) %></p>
    <% End If %>
    <form method="post" action="/modules/solicitudes/alta_do.asp" class="formulario-vertical">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(token) %>">
        <label>MSISDN (opcional)
            <input type="text" name="msisdn" maxlength="9">
        </label>
        <% If esInclusion Then %>
        <label>Numero de denuncia
            <input type="text" name="numDenuncia" maxlength="20" required>
        </label>
        <label>Fecha del robo
            <input type="date" name="fechaRobo" required>
        </label>
        <label>Fecha de la denuncia
            <input type="date" name="fechaDenuncia" required>
        </label>
        <% Else %>
        <label>Numero de denuncia (opcional)
            <input type="text" name="numDenuncia" maxlength="20">
        </label>
        <label>Fecha del robo (opcional)
            <input type="date" name="fechaRobo">
        </label>
        <label>Fecha de la denuncia (opcional)
            <input type="date" name="fechaDenuncia">
        </label>
        <% End If %>
        <label>NIF/CIF del cliente
            <input type="text" name="identCliente" maxlength="20" value="<%= Server.HTMLEncode(Session("Sol_PrefillIdent")) %>"<%= atributoReadonly %> required>
        </label>
        <label>Nombre
            <input type="text" name="nombreCliente" maxlength="200" value="<%= Server.HTMLEncode(Session("Sol_PrefillNombre")) %>"<%= atributoReadonly %> required>
        </label>
        <label>Primer apellido
            <input type="text" name="apellido1Cliente" maxlength="200" value="<%= Server.HTMLEncode(Session("Sol_PrefillApellido1")) %>"<%= atributoReadonly %> required>
        </label>
        <label>Segundo apellido
            <input type="text" name="apellido2Cliente" maxlength="200" value="<%= Server.HTMLEncode(Session("Sol_PrefillApellido2")) %>"<%= atributoReadonly %>>
        </label>
        <button type="submit">Crear solicitud</button>
    </form>
</main>
</div>
<!--#include virtual="/include/layout/footer.asp"-->
</div>
</body>
</html>
