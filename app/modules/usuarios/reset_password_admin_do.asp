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

If Not IsNumeric(Request.Form("cod")) Then
    conn.Close
    Response.Redirect "/errors/error.asp"
    Response.End
End If

Dim codIndiceUsuario, passwordNueva, passwordNueva2, urlVolver
codIndiceUsuario = CLng(Request.Form("cod"))
passwordNueva = Request.Form("passwordNueva")
passwordNueva2 = Request.Form("passwordNueva2")
urlVolver = "/modules/usuarios/reset_password_admin.asp?cod=" & codIndiceUsuario & "&error="

If Len(passwordNueva) < 10 Then
    conn.Close
    Response.Redirect urlVolver & Server.URLEncode("La nueva contrasena debe tener al menos 10 caracteres.")
    Response.End
ElseIf passwordNueva <> passwordNueva2 Then
    conn.Close
    Response.Redirect urlVolver & Server.URLEncode("Las dos contrasenas nuevas no coinciden.")
    Response.End
End If

Dim u
Set u = ObtenerUsuarioPorIndice(conn, codIndiceUsuario)
If u Is Nothing Then
    conn.Close
    Response.Redirect "/errors/error.asp"
    Response.End
End If

Dim maxHistorico, diasCaducidad
maxHistorico = ValorParametroEntero(conn, "NUM_OLD_PASSWORD", 5)
diasCaducidad = ValorParametroEntero(conn, "DURACION_PASSWORD", 60)

Dim resultado
resultado = CambiarPasswordConHistorico(conn, codIndiceUsuario, u("Password"), u("AlgoritmoPassword"), passwordNueva, maxHistorico, diasCaducidad)

If resultado <> "" Then
    conn.Close
    Set conn = Nothing
    Response.Redirect urlVolver & Server.URLEncode(resultado)
    Response.End
End If

Call RegistrarAuditoria(conn, Session("CodIndiceUsuario"), "CAMBIO_PASSWORD_ADMIN", "USUARIOS", codIndiceUsuario, "OK", "Reset realizado por perfil Seguridad")

conn.Close
Set conn = Nothing

Call EstablecerMensajeFlash("exito", "Contrasena del usuario actualizada correctamente.")
Response.Redirect "/modules/usuarios/listar.asp"
%>
