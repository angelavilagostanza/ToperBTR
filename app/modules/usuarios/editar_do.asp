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

Dim codIndiceUsuario, nombre, apellidos, codPerfil, urlVolver
codIndiceUsuario = CLng(Request.Form("cod"))
nombre = Trim(Request.Form("nombre"))
apellidos = Trim(Request.Form("apellidos"))
codPerfil = Request.Form("perfil")
urlVolver = "/modules/usuarios/editar.asp?cod=" & codIndiceUsuario & "&error="

Dim errorMsg
errorMsg = ""

If Not EsCadenaNoVacia(nombre) Or Not EsCadenaNoVacia(apellidos) Then
    errorMsg = "Nombre y apellidos son obligatorios."
ElseIf Not IsNumeric(codPerfil) Then
    errorMsg = "Debe seleccionar un perfil."
End If

If errorMsg <> "" Then
    conn.Close
    Set conn = Nothing
    Response.Redirect urlVolver & Server.URLEncode(errorMsg)
    Response.End
End If

Dim u
Set u = ObtenerUsuarioPorIndice(conn, codIndiceUsuario)
If u Is Nothing Then
    conn.Close
    Response.Redirect "/errors/error.asp"
    Response.End
End If

Call ActualizarDatosUsuario(conn, codIndiceUsuario, nombre, apellidos, CLng(codPerfil))

Call RegistrarAuditoria(conn, Session("CodIndiceUsuario"), "USUARIO_MODIFICACION", "USUARIOS", codIndiceUsuario, "OK", "Nombre/apellidos/perfil actualizados")

conn.Close
Set conn = Nothing

Call EstablecerMensajeFlash("exito", "Usuario modificado correctamente.")
Response.Redirect "/modules/usuarios/listar.asp"
Response.End
%>
