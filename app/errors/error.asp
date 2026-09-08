<%
Dim msg, texto
msg = Request.QueryString("msg")
Select Case msg
    Case "sinpermiso"
        texto = "No tiene permisos para acceder a esta seccion."
    Case "csrf"
        texto = "La sesion del formulario ha expirado. Vuelva a intentarlo."
    Case "bd"
        texto = "No se ha podido conectar con la base de datos. Intentelo de nuevo en unos minutos o contacte con el administrador si el problema persiste."
    Case "consulta"
        texto = "Se ha producido un error al acceder a los datos. Intentelo de nuevo mas tarde."
    Case "ficherosindato"
        texto = "No se ha indicado ningun fichero a descargar."
    Case "ficheronoencontrado"
        texto = "El fichero solicitado no existe o ya no esta disponible."
    Case "imeinoyoigo"
        texto = "El IMEI introducido no pertenece a un cliente activo de Grupo MASMOVIL. Solo se pueden bloquear IMEIs de clientes Grupo MASMOVIL."
    Case "imeibloqueado"
        texto = "El IMEI introducido ya se encuentra bloqueado. No es posible volver a bloquearlo."
    Case "noencontrado"
        texto = "La pagina solicitada no existe o ha sido eliminada."
    Case Else
        texto = "Se ha producido un error. Intentelo de nuevo mas tarde."
End Select
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Aviso</title>
<link rel="stylesheet" href="/assets/css/site.css">
</head>
<body class="pagina-error">
<div class="caja-error">
<h1>Aviso</h1>
<p><%= Server.HTMLEncode(texto) %></p>
<p><a href="/default.asp">Volver al inicio</a></p>
</div>
</body>
</html>
