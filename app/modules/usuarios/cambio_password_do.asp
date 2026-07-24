<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/usuarios_dal.asp"-->
<%
Call RequiereSesion()

If Not TokenCsrfValido(Request.Form("csrf")) Then
    Response.Redirect "/errors/error.asp?msg=csrf"
    Response.End
End If

Dim passwordActual, passwordNueva, passwordNueva2
passwordActual = Request.Form("passwordActual")
passwordNueva = Request.Form("passwordNueva")
passwordNueva2 = Request.Form("passwordNueva2")

If Len(passwordNueva) < 10 Then
    Response.Redirect "/modules/usuarios/cambio_password.asp?error=" & Server.URLEncode("La nueva contrasena debe tener al menos 10 caracteres.")
    Response.End
ElseIf passwordNueva <> passwordNueva2 Then
    Response.Redirect "/modules/usuarios/cambio_password.asp?error=" & Server.URLEncode("Las dos contrasenas nuevas no coinciden.")
    Response.End
End If

Dim conn
Set conn = NuevaConexion()

Dim u
Set u = ObtenerUsuarioPorIndice(conn, Session("CodIndiceUsuario"))

If u Is Nothing Then
    conn.Close
    Response.Redirect "/errors/error.asp"
    Response.End
End If

If Not VerificarPassword(passwordActual, u("Password"), u("AlgoritmoPassword")) Then
    conn.Close
    Set conn = Nothing
    Response.Redirect "/modules/usuarios/cambio_password.asp?error=" & Server.URLEncode("La contrasena actual no es correcta.")
    Response.End
End If

Dim maxHistorico, diasCaducidad
maxHistorico = ValorParametroEntero(conn, "NUM_OLD_PASSWORD", 5)
diasCaducidad = ValorParametroEntero(conn, "DURACION_PASSWORD", 60)

Dim resultado
resultado = CambiarPasswordConHistorico(conn, Session("CodIndiceUsuario"), u("Password"), u("AlgoritmoPassword"), passwordNueva, maxHistorico, diasCaducidad)

If resultado <> "" Then
    conn.Close
    Set conn = Nothing
    Response.Redirect "/modules/usuarios/cambio_password.asp?error=" & Server.URLEncode(resultado)
    Response.End
End If

Call RegistrarAuditoria(conn, Session("CodIndiceUsuario"), "CAMBIO_PASSWORD", "USUARIOS", Session("CodIndiceUsuario"), "OK", "Autoservicio")

conn.Close
Set conn = Nothing

Call EstablecerMensajeFlash("exito", "Contrasena actualizada correctamente.")
Response.Redirect "/modules/home/home.asp"
%>
