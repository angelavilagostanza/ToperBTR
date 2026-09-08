<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/bloqueos_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()

Dim codAdmin, codTramitacion
codAdmin = CodigoPerfil(conn, "Administrador")
codTramitacion = CodigoPerfil(conn, "Tramitación")
Call RequierePerfil(Array(codAdmin, codTramitacion))

If Not TokenCsrfValido(Request.Form("csrf")) Then
    conn.Close
    Response.Redirect "/errors/error.asp?msg=csrf"
    Response.End
End If

' Request.Form con nombre multiple devuelve valores separados por coma en ASP clasico
Dim seleccionRaw
seleccionRaw = Request.Form("seleccion")

If Not EsCadenaNoVacia(seleccionRaw) Then
    conn.Close
    Set conn = Nothing
    Call EstablecerMensajeFlash("aviso", "No se ha seleccionado ningun IMEI para desbloquear.")
    Response.Redirect "/modules/bloqueos/bloqueo.asp"
    Response.End
End If

Dim imeisArray
imeisArray = Split(seleccionRaw, ",")

' Sanear cada IMEI: quitar espacios y verificar formato antes de pasar al DAL.
' El DAL filtra los que no midan 15 chars, pero rechazar aqui evita queries innecesarias.
Dim i, imeisValidos(), numValidos
numValidos = 0
ReDim imeisValidos(UBound(imeisArray))
For i = 0 To UBound(imeisArray)
    Dim imeiCandidato
    imeiCandidato = Trim(imeisArray(i))
    If Len(imeiCandidato) = 15 And CumpleFormato(imeiCandidato, "^\d{15}$") Then
        imeisValidos(numValidos) = imeiCandidato
        numValidos = numValidos + 1
    End If
Next

If numValidos = 0 Then
    conn.Close
    Set conn = Nothing
    Call EstablecerMensajeFlash("aviso", "Ninguno de los IMEIs seleccionados tiene formato valido.")
    Response.Redirect "/modules/bloqueos/bloqueo.asp"
    Response.End
End If

ReDim Preserve imeisValidos(numValidos - 1)

Dim desbloqueados
desbloqueados = DesbloquearIMEIsLote(conn, imeisValidos)

' Auditar cada IMEI que se haya desbloqueado con exito.
' DesbloquearIMEIsLote solo confirma los que ExcluirIMEI devolvio True, pero no sabemos
' cuales especificamente. Se audita como lote indicando el total gestionado.
Dim detalle
detalle = "total_solicitado=" & numValidos & ";desbloqueados=" & desbloqueados
Call RegistrarAuditoria(conn, Session("CodIndiceUsuario"), "BLOQUEO_QUITAR", "LISTA_NEGRA", 0, "OK", detalle)

conn.Close
Set conn = Nothing

If desbloqueados = numValidos Then
    Call EstablecerMensajeFlash("exito", "Se ha desbloqueado correctamente " & desbloqueados & " IMEI(s).")
ElseIf desbloqueados > 0 Then
    Call EstablecerMensajeFlash("aviso", "Se han desbloqueado " & desbloqueados & " de " & numValidos & " IMEI(s). Los restantes podrian no tener bloqueo directo activo.")
Else
    Call EstablecerMensajeFlash("error", "No se ha podido desbloquear ningun IMEI. Es posible que no tengan bloqueo directo activo.")
End If

Response.Redirect "/modules/bloqueos/bloqueo.asp"
Response.End
%>
