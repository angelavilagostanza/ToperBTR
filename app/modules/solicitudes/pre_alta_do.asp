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

Dim imei, tipoClave
imei = Trim(Request.Form("imei"))
tipoClave = Request.Form("tipo")

If Not CumpleFormato(imei, "^\d{15}$") Then
    conn.Close
    Set conn = Nothing
    Response.Redirect "/modules/solicitudes/pre_alta.asp?error=" & Server.URLEncode("El IMEI debe tener 15 digitos numericos.")
    Response.End
End If

If tipoClave <> "I" And tipoClave <> "E" Then
    conn.Close
    Set conn = Nothing
    Response.Redirect "/modules/solicitudes/pre_alta.asp?error=" & Server.URLEncode("Debe seleccionar un tipo de solicitud.")
    Response.End
End If

If ImeiEnCurso(conn, imei) Then
    conn.Close
    Set conn = Nothing
    Response.Redirect "/modules/solicitudes/pre_alta.asp?error=" & Server.URLEncode("Ya existe una solicitud en curso para este IMEI.")
    Response.End
End If

If ImeiEnListaNegra(conn, imei) Then
    conn.Close
    Set conn = Nothing

    Response.Redirect "/modules/solicitudes/pre_alta.asp?error=" & _
        Server.URLEncode("El IMEI ya se encuentra en lista negra.")

    Response.End
End If

Session("Sol_Imei") = imei
Session("Sol_TipoClave") = tipoClave
Session("Sol_PrefillNombre") = ""
Session("Sol_PrefillApellido1") = ""
Session("Sol_PrefillApellido2") = ""
Session("Sol_PrefillIdent") = ""
Session("Sol_PrefillBloqueado") = False

If tipoClave = "E" Then
    Dim padre
    Set padre = BuscarSolicitudPadrePorImei(conn, imei)
    If padre Is Nothing Then
        conn.Close
        Set conn = Nothing
        Response.Redirect "/modules/solicitudes/pre_alta.asp?error=" & Server.URLEncode("No existe una solicitud de Inclusion previa para este IMEI; no se puede crear una Exclusion.")
        Response.End
    End If
    Session("Sol_PrefillNombre") = padre("NombreCliente")
    Session("Sol_PrefillApellido1") = padre("Apellido1Cliente")
    Session("Sol_PrefillApellido2") = padre("Apellido2Cliente")
    Session("Sol_PrefillIdent") = padre("IdentCliente")
    'Session("Sol_PrefillBloqueado") = True
	Session("Sol_PrefillBloqueado") = False
End If

conn.Close
Set conn = Nothing
Response.Redirect "/modules/solicitudes/alta.asp"
Response.End
%>
