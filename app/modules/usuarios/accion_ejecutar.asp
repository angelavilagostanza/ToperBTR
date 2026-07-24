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

Dim accion, seleccionCsv, listaIdsStr, listaIdsLng, i
accion = Request.Form("accion")
seleccionCsv = Request.Form("seleccion")

If Not EsCadenaNoVacia(seleccionCsv) Then
    conn.Close
    Response.Redirect "/modules/usuarios/listar.asp"
    Response.End
End If

listaIdsStr = Split(seleccionCsv, ",")
ReDim listaIdsLng(UBound(listaIdsStr))
For i = 0 To UBound(listaIdsStr)
    listaIdsLng(i) = CLng(listaIdsStr(i))
Next

Dim ok, codBorrado, codActivo
codBorrado = CodigoEstadoUsuario(conn, "Borrado")
codActivo = CodigoEstadoUsuario(conn, "Activo")

If accion = "baja" Then
    ok = DarBajaUsuariosLote(conn, listaIdsLng, codBorrado)
Else
    ok = DesbloquearUsuariosLote(conn, listaIdsLng, codActivo)
End If

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
Response.Redirect "/modules/usuarios/listar.asp"
%>
