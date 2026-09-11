<!--#include virtual="/include/bootstrap.asp"-->
<%
' UTILIDAD TEMPORAL - ELIMINAR DESPUES DE USAR
' Resetea la password de un usuario a un valor conocido usando las iteraciones actuales.
' Solo accesible desde localhost para seguridad minima.

If Request.ServerVariables("REMOTE_ADDR") <> "127.0.0.1" And _
   Request.ServerVariables("REMOTE_ADDR") <> "10.27.47.124" Then
    Response.Write "Solo accesible desde el servidor."
    Response.End
End If

Dim accion, usuario, nuevaPass, nuevoHash, conn, rs
accion    = Request.Form("accion")
usuario   = Trim(Request.Form("usuario"))
nuevaPass = Request.Form("nuevapass")
%>
<!DOCTYPE html>
<html lang="es">
<head><meta charset="utf-8"><title>Reset Hash - ToperBTR</title>
<link rel="stylesheet" href="/assets/css/site.css">
<link rel="icon" href="/images/ToperBTR_favicon.ico" type="image/x-icon">
</head>
<body class="pagina-error">
<div class="caja-error" style="min-width:420px">
<h1>Reset Hash de Contrasena</h1>
<p style="color:#8a1c0a;font-weight:600">UTILIDAD TEMPORAL - ELIMINAR DESPUES DE USAR</p>
<p>Iteraciones actuales: <strong><%= Application("PBKDF2Iteraciones") %></strong></p>

<% If accion = "reset" And EsCadenaNoVacia(usuario) And EsCadenaNoVacia(nuevaPass) Then
    nuevoHash = CrearHashPbkdf2(nuevaPass, CLng(Application("PBKDF2Iteraciones")))
    Set conn = NuevaConexion()
    Set rs = EjecutarConsulta(conn, "SELECT COD_INDICE_USUARIO FROM USUARIOS WHERE COD_USUARIO = ?", Array(usuario))
    If rs.EOF Then
        rs.Close : conn.Close : Set conn = Nothing
%>
    <p style="color:#8a1c0a">Usuario '<%= Server.HTMLEncode(usuario) %>' no encontrado.</p>
<% Else
        Dim codIdx : codIdx = rs("COD_INDICE_USUARIO")
        rs.Close
        Call EjecutarNonQuery(conn, "UPDATE USUARIOS SET PASSWORD = ?, ALGORITMO_PASSWORD = 'PBKDF2SHA256', NUM_FALLOS = 0 WHERE COD_INDICE_USUARIO = ?", Array(nuevoHash, codIdx))
        conn.Close : Set conn = Nothing
%>
    <p style="color:#1e6b34;font-weight:600">OK: contrasena de '<%= Server.HTMLEncode(usuario) %>' reseteada a '<%= Server.HTMLEncode(nuevaPass) %>' con <%= Application("PBKDF2Iteraciones") %> iteraciones.</p>
<% End If
End If %>

<form method="post">
    <input type="hidden" name="accion" value="reset">
    <label>Usuario (COD_USUARIO):<br>
        <input type="text" name="usuario" value="<%= Server.HTMLEncode(usuario) %>" style="width:100%;padding:.4rem;margin:.3rem 0">
    </label>
    <label>Nueva contrasena:<br>
        <input type="password" name="nuevapass" style="width:100%;padding:.4rem;margin:.3rem 0">
    </label>
    <button type="submit" style="margin-top:.8rem;width:100%;padding:.6rem;background:#e0801a;color:#fff;border:none;border-radius:4px;font-weight:600;cursor:pointer">
        Resetear hash
    </button>
</form>

<p style="margin-top:1.5rem"><a href="/modules/login/login.asp">Ir al login</a></p>
</div>
</body>
</html>
