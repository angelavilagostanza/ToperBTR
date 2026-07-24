<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/solicitudes_dal.asp"-->
<%
Call RequiereSesion()

If Not EsCadenaNoVacia(Session("Sol_Imei")) Then
    Response.Redirect "/modules/solicitudes/pre_alta.asp"
    Response.End
End If

If Not TokenCsrfValido(Request.Form("csrf")) Then
    Response.Redirect "/errors/error.asp?msg=csrf"
    Response.End
End If

Dim imei, tipoClave, esInclusion
imei = Session("Sol_Imei")
tipoClave = Session("Sol_TipoClave")
esInclusion = (tipoClave = "I")

Dim msisdn, numDenuncia, fechaRobo, fechaDenuncia, identCliente, nombreCliente, apellido1Cliente, apellido2Cliente
msisdn = Trim(Request.Form("msisdn"))
numDenuncia = Trim(Request.Form("numDenuncia"))
fechaRobo = Trim(Request.Form("fechaRobo"))
fechaDenuncia = Trim(Request.Form("fechaDenuncia"))
identCliente = Trim(Request.Form("identCliente"))
nombreCliente = Trim(Request.Form("nombreCliente"))
apellido1Cliente = Trim(Request.Form("apellido1Cliente"))
apellido2Cliente = Trim(Request.Form("apellido2Cliente"))

Dim errorMsg
errorMsg = ""

If EsCadenaNoVacia(msisdn) And Not CumpleFormato(msisdn, "^\d{9}$") Then
    errorMsg = "El MSISDN, si se indica, debe tener 9 digitos numericos."
ElseIf Not EsCadenaNoVacia(identCliente) Or Not EsCadenaNoVacia(nombreCliente) Or Not EsCadenaNoVacia(apellido1Cliente) Then
    errorMsg = "NIF/CIF, nombre y primer apellido del cliente son obligatorios."
ElseIf esInclusion And Not EsCadenaNoVacia(numDenuncia) Then
    errorMsg = "El numero de denuncia es obligatorio para una Inclusion."
ElseIf esInclusion And Not IsDate(fechaRobo) Then
    errorMsg = "La fecha del robo es obligatoria y valida para una Inclusion."
ElseIf esInclusion And Not IsDate(fechaDenuncia) Then
    errorMsg = "La fecha de la denuncia es obligatoria y valida para una Inclusion."
ElseIf EsCadenaNoVacia(fechaRobo) And Not IsDate(fechaRobo) Then
    errorMsg = "La fecha del robo no es valida."
ElseIf EsCadenaNoVacia(fechaDenuncia) And Not IsDate(fechaDenuncia) Then
    errorMsg = "La fecha de la denuncia no es valida."
End If

If errorMsg <> "" Then
    Response.Redirect "/modules/solicitudes/alta.asp?error=" & Server.URLEncode(errorMsg)
    Response.End
End If

Dim conn
Set conn = NuevaConexion()
Call RequiereAccion(conn, CodigoAccion(conn, "ESCRITURA_SOLICITUDES"))

If ImeiEnCurso(conn, imei) Then
    conn.Close
    Set conn = Nothing
    Response.Redirect "/modules/solicitudes/pre_alta.asp?error=" & Server.URLEncode("Ya existe una solicitud en curso para este IMEI.")
    Response.End
End If

Dim fechaRoboValor, fechaDenunciaValor, msisdnValor, numDenunciaValor
If IsDate(fechaRobo) Then fechaRoboValor = CDate(fechaRobo) Else fechaRoboValor = Null
If IsDate(fechaDenuncia) Then fechaDenunciaValor = CDate(fechaDenuncia) Else fechaDenunciaValor = Null
If EsCadenaNoVacia(msisdn) Then msisdnValor = msisdn Else msisdnValor = Null
If EsCadenaNoVacia(numDenuncia) Then numDenunciaValor = numDenuncia Else numDenunciaValor = Null

Dim resultado
Set resultado = CrearSolicitudCompleta(conn, tipoClave, imei, msisdnValor, numDenunciaValor, fechaRoboValor, fechaDenunciaValor, _
    identCliente, nombreCliente, apellido1Cliente, apellido2Cliente, Session("CodIndiceUsuario"), Session("CodOperador"))

If Not resultado("Exito") Then
    conn.Close
    Set conn = Nothing
    Response.Redirect "/modules/solicitudes/alta.asp?error=" & Server.URLEncode(resultado("Mensaje"))
    Response.End
End If

Call RegistrarAuditoria(conn, Session("CodIndiceUsuario"), "SOLICITUD_ALTA", "SOLICITUD", resultado("CodIndiceSol"), "OK", "Codigo: " & resultado("CodSolicitud"))

conn.Close
Set conn = Nothing

Session("Sol_Imei") = ""
Session("Sol_TipoClave") = ""

Call EstablecerMensajeFlash("exito", "Solicitud " & resultado("CodSolicitud") & " creada correctamente.")
Response.Redirect "/modules/solicitudes/detalle.asp?cod=" & resultado("CodIndiceSol")
%>
