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

Dim accion, seleccionCsv, volver, listaIdsStr, listaIdsLng, i
accion = Request.Form("accion")
seleccionCsv = Request.Form("seleccion")

' Ver el mismo comentario en confirmar_accion.asp: ruta validada contra open-redirect.
volver = Request.Form("volver")
If Not (EsCadenaNoVacia(volver) And Left(volver, Len("/modules/usuarios/listar.asp")) = "/modules/usuarios/listar.asp") Then
    volver = "/modules/usuarios/listar.asp"
End If

If Not EsCadenaNoVacia(seleccionCsv) Then
    conn.Close
    Response.Redirect volver
    Response.End
End If

listaIdsStr = Split(seleccionCsv, ",")
ReDim listaIdsLng(UBound(listaIdsStr))
For i = 0 To UBound(listaIdsStr)
    listaIdsLng(i) = CLng(listaIdsStr(i))
Next

Dim ok, codBorrado, codActivo, codBloqueado
codBorrado = CodigoEstadoUsuario(conn, "Borrado")
codActivo = CodigoEstadoUsuario(conn, "Activo")
codBloqueado = CodigoEstadoUsuario(conn, "Bloqueado")

Select Case accion
    Case "baja"
        ok = DarBajaUsuariosLote(conn, listaIdsLng, codBorrado)
    Case "bloqueo"
        ok = BloquearUsuariosLote(conn, listaIdsLng, codBloqueado, codBorrado)
    Case "desbloqueo"
        ok = DesbloquearUsuariosLote(conn, listaIdsLng, codActivo, codBorrado)
    Case Else
        conn.Close
        Response.Redirect volver
        Response.End
End Select

If ok Then
    For i = 0 To UBound(listaIdsLng)
        Call RegistrarAuditoria(conn, Session("CodIndiceUsuario"), "USUARIO_" & UCase(accion), "USUARIOS", listaIdsLng(i), "OK", "")
    Next
    Call EstablecerMensajeFlash("exito", "Operacion realizada correctamente.")
Else
    Call EstablecerMensajeFlash("error", MensajeErrorGenerico())
End If

conn.Close
Set conn = Nothing
Response.Redirect volver
%>
