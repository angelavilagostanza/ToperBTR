<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/solicitudes_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequiereAccion(conn, CodigoAccion(conn, "ESCRITURA_SOLICITUDES"))

If Not TokenCsrfValido(Request.Form("csrf")) Then
    conn.Close
    Response.Redirect "/errors/error.asp?msg=csrf"
    Response.End
End If

Dim seleccionCsv
seleccionCsv = Request.Form("seleccion")

If Not EsCadenaNoVacia(seleccionCsv) Then
    conn.Close
    Set conn = Nothing
    Response.Redirect "/modules/solicitudes/consulta.asp"
    Response.End
End If

Dim listaIdsStr, listaIdsLng, i
listaIdsStr = Split(seleccionCsv, ",")
ReDim listaIdsLng(UBound(listaIdsStr))
For i = 0 To UBound(listaIdsStr)
    listaIdsLng(i) = CLng(listaIdsStr(i))
Next

Dim codCancelada
codCancelada = CodigoEstadoSolicitud(conn, "Cancelada")

Dim ok
ok = CancelarSolicitudesLote(conn, listaIdsLng, codCancelada)

If ok Then
    For i = 0 To UBound(listaIdsLng)
        Call RegistrarAuditoria(conn, Session("CodIndiceUsuario"), "SOLICITUD_CANCELAR", "SOLICITUD", listaIdsLng(i), "OK", "")
    Next
    Call EstablecerMensajeFlash("exito", "Solicitud(es) cancelada(s) correctamente.")
Else
    Call EstablecerMensajeFlash("error", MensajeErrorGenerico())
End If

conn.Close
Set conn = Nothing
Response.Redirect "/modules/solicitudes/consulta.asp"
%>
