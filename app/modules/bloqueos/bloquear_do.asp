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

Dim imei, codRazon, comentario
imei       = Trim(Request.Form("imei") & "")
codRazon   = Trim(Request.Form("razon") & "")
comentario = Trim(Request.Form("comentario") & "")

' Validaciones — corrigen el bug real de BusquedaBloqueoForm.validate() del legacy donde
' la comprobacion numerica del IMEI era codigo muerto (condicion duplicada en else if).
If Not EsCadenaNoVacia(imei) Then
    conn.Close
    Call EstablecerMensajeFlash("error", "Debe introducir un IMEI.")
    Response.Redirect "/modules/bloqueos/bloquear.asp"
    Response.End
End If

If Len(imei) <> 15 Or Not CumpleFormato(imei, "^\d{15}$") Then
    conn.Close
    Call EstablecerMensajeFlash("error", "El IMEI debe tener exactamente 15 digitos numericos.")
    Response.Redirect "/modules/bloqueos/bloquear.asp"
    Response.End
End If

If Not EsCadenaNoVacia(codRazon) Or Not IsNumeric(codRazon) Then
    conn.Close
    Call EstablecerMensajeFlash("error", "Debe seleccionar una razon de bloqueo.")
    Response.Redirect "/modules/bloqueos/bloquear.asp"
    Response.End
End If

Dim resultado
resultado = InsertarBloqueoDirecto(conn, imei, CLng(codRazon), comentario)

Select Case resultado
    Case "OK"
        Call RegistrarAuditoria(conn, Session("CodIndiceUsuario"), "BLOQUEO_INSERTAR", "LISTA_NEGRA", imei, "OK", "razon=" & codRazon)
        Call EstablecerMensajeFlash("exito", "El IMEI " & imei & " ha sido bloqueado correctamente.")
        conn.Close
        Set conn = Nothing
        Response.Redirect "/modules/bloqueos/bloqueo.asp"
        Response.End

    Case "BLOQUEADO"
        Call RegistrarAuditoria(conn, Session("CodIndiceUsuario"), "BLOQUEO_INSERTAR", "LISTA_NEGRA", imei, "YA_BLOQUEADO", "")
        conn.Close
        Set conn = Nothing
        Response.Redirect "/errors/error.asp?msg=imeibloqueado"
        Response.End

    Case "NO_YOIGO"
        Call RegistrarAuditoria(conn, Session("CodIndiceUsuario"), "BLOQUEO_INSERTAR", "LISTA_NEGRA", imei, "NO_YOIGO", "")
        conn.Close
        Set conn = Nothing
        Response.Redirect "/errors/error.asp?msg=imeinoyoigo"
        Response.End

    Case Else
        Call EstablecerMensajeFlash("error", "No se ha podido realizar el bloqueo. Intentelo de nuevo.")
        conn.Close
        Set conn = Nothing
        Response.Redirect "/modules/bloqueos/bloquear.asp"
        Response.End
End Select
%>
