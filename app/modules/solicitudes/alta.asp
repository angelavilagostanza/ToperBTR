<!--#include virtual="/include/bootstrap.asp"-->
<%
' Hallazgo B2: hasta ahora el unico permiso comprobado aqui era RequiereSesion() (estar
' logueado). En la practica esta pagina solo pinta algo util si Session("Sol_Imei") tiene
' valor, y esa variable SOLO la fija pre_alta_do.asp, que si tiene el guard correcto
' (RequiereAccion ESCRITURA_SOLICITUDES) - asi que hoy es imposible llegar aqui con datos
' utiles sin haber pasado por ese control. Pero es una dependencia fragil, no una garantia
' explicita: si en el futuro algo mas llegara a fijar esa variable de sesion sin pasar por
' ese control, aqui no habria nada que lo impidiera. Se añade el mismo guard explicito que
' ya tienen pre_alta.asp/pre_alta_do.asp, por consistencia con el principio ya aplicado en
' todo el proyecto (docs/04-matriz-permisos.md: "cada operacion sensible comprueba su
' permiso en el servidor, nunca solo por ocultacion/indireccion").
Dim conn
Set conn = NuevaConexion()
Call RequiereAccion(conn, CodigoAccion(conn, "ESCRITURA_SOLICITUDES"))
conn.Close
Set conn = Nothing

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

    <nav class="modulo-nav">
        <a href="/modules/solicitudes/consulta.asp" class="modulo-nav-item">Buscar</a>
        <span class="modulo-nav-item activo">+ Nueva solicitud</span>
    </nav>

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
            <input type="date" maxlength="50" name="fechaRobo" required>
        </label>
        <label>Fecha de la denuncia
            <input type="date" maxlength="50" name="fechaDenuncia" required>
        </label>
        <% Else %>
        <label>Numero de denuncia (opcional)
            <input type="text" name="numDenuncia" maxlength="20">
        </label>
        <label>Fecha del robo (opcional)
            <input type="date" maxlength="50" name="fechaRobo">
        </label>
        <label>Fecha de la denuncia (opcional)
            <input type="date" maxlength="50" name="fechaDenuncia">
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
