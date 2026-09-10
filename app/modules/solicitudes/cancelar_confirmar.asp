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
    Call EstablecerMensajeFlash("error", "No se ha seleccionado ninguna solicitud.")
    Response.Redirect "/modules/solicitudes/consulta.asp"
    Response.End
End If

Dim listaIdsRaw, i
listaIdsRaw = Split(seleccionCsv, ",")

' Validar: solo IDs numericos y cancelables. Un ID no cancelable en este punto
' indica manipulacion del request (checkbox deshabilitado habilitado manualmente).
Dim listaIdsVal(), nVal
ReDim listaIdsVal(UBound(listaIdsRaw))
nVal = 0
For i = 0 To UBound(listaIdsRaw)
    Dim rawId
    rawId = Trim(listaIdsRaw(i))
    If Not IsNumeric(rawId) Then
        conn.Close
        Response.Redirect "/errors/error.asp?msg=solicitud_invalida"
        Response.End
    End If
    Dim idLng
    idLng = CLng(rawId)
    If Not EsSolicitudCancelable(conn, idLng) Then
        conn.Close
        Call EstablecerMensajeFlash("error", "Una o mas solicitudes seleccionadas no pueden cancelarse.")
        Response.Redirect "/modules/solicitudes/consulta.asp"
        Response.End
    End If
    listaIdsVal(nVal) = idLng
    nVal = nVal + 1
Next
If nVal = 0 Then
    conn.Close
    Call EstablecerMensajeFlash("error", "No se ha seleccionado ninguna solicitud valida.")
    Response.Redirect "/modules/solicitudes/consulta.asp"
    Response.End
End If
ReDim Preserve listaIdsVal(nVal - 1)

Dim tokenConfirm
tokenConfirm = GenerarTokenCsrf()

' Reconstruir CSV solo con los IDs validados
Dim seleccionValidada, j
seleccionValidada = ""
For j = 0 To UBound(listaIdsVal)
    If j > 0 Then seleccionValidada = seleccionValidada & ","
    seleccionValidada = seleccionValidada & CStr(listaIdsVal(j))
Next
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Confirmar cancelacion</title>
<link rel="stylesheet" href="/assets/css/site.css">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Confirmar cancelacion de solicitud(es)</h2>
    <ul>
    <% For j = 0 To UBound(listaIdsVal)
        Dim s
        Set s = ObtenerSolicitudPorIndice(conn, listaIdsVal(j))
        If Not s Is Nothing Then %>
        <li><%= Server.HTMLEncode(s("CodSolicitud")) %> &mdash; IMEI <%= Server.HTMLEncode(s("Imei")) %></li>
        <% End If
    Next %>
    </ul>
    <form method="post" action="/modules/solicitudes/cancelar_ejecutar.asp">
        <input type="hidden" name="csrf" value="<%= Server.HTMLEncode(tokenConfirm) %>">
        <input type="hidden" name="seleccion" value="<%= Server.HTMLEncode(seleccionValidada) %>">
        <button type="submit">Confirmar cancelacion</button>
        <a href="/modules/solicitudes/consulta.asp">Cancelar</a>
    </form>
</main>
</div>
<!--#include virtual="/include/layout/footer.asp"-->
</div>
</body>
</html>
<%
conn.Close
Set conn = Nothing
%>
