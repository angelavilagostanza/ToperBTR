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

' Segunda validacion (defensa en profundidad): cada ID debe ser numerico y
' cancelable. cancelar_confirmar.asp ya lo valido, pero un POST directo o una
' solicitud que cambio de estado entre pasos podria llegar aqui igualmente.
ReDim listaIdsLng(UBound(listaIdsStr))
For i = 0 To UBound(listaIdsStr)
    Dim rawId
    rawId = Trim(listaIdsStr(i))
    If Not IsNumeric(rawId) Then
        conn.Close
        Response.Redirect "/errors/error.asp?msg=solicitud_invalida"
        Response.End
    End If
    listaIdsLng(i) = CLng(rawId)
    If Not EsSolicitudCancelable(conn, listaIdsLng(i)) Then
        conn.Close
        Call EstablecerMensajeFlash("error", "Una o mas solicitudes ya no pueden cancelarse. Revisa el estado actual y vuelve a intentarlo.")
        Response.Redirect "/modules/solicitudes/consulta.asp"
        Response.End
    End If
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
