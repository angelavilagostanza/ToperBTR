<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/incidencias_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequiereAccion(conn, CodigoAccion(conn, "ESCRITURA_INCIDENCIAS"))

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

Dim codIndIncidencia, comentario
codIndIncidencia = CLng(Request.Form("cod"))
comentario = Trim(Request.Form("comentario"))

If Not EsCadenaNoVacia(comentario) Then
    conn.Close
    Set conn = Nothing
    Response.Redirect "/modules/incidencias/detalle.asp?cod=" & codIndIncidencia & "&error=" & Server.URLEncode("El comentario no puede estar vacio.")
    Response.End
End If

If Len(comentario) > 500 Then
    comentario = Left(comentario, 500)
End If

Call InsertarComentario(conn, codIndIncidencia, Session("CodIndiceUsuario"), comentario)
Call RegistrarAuditoria(conn, Session("CodIndiceUsuario"), "INCIDENCIA_COMENTARIO", "INCIDENCIA", codIndIncidencia, "OK", "")

conn.Close
Set conn = Nothing

Call EstablecerMensajeFlash("exito", "Comentario anadido correctamente.")
Response.Redirect "/modules/incidencias/detalle.asp?cod=" & codIndIncidencia
Response.End
%>
