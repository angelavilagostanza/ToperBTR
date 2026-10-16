<!--#include virtual="/include/bootstrap.asp"-->
<%
If Session("Autenticado") = True Then
    Dim connLogout
    Set connLogout = NuevaConexion()
    Call RegistrarAuditoria(connLogout, Session("CodIndiceUsuario"), "LOGOUT", "USUARIOS", Session("CodIndiceUsuario"), "OK", "")
    connLogout.Close
    Set connLogout = Nothing
End If
Session.Abandon
Response.AddHeader "Set-Cookie", "ToperBTRFp=; Path=/; Expires=Thu, 01 Jan 1970 00:00:00 GMT"
Response.Redirect "/modules/login/login.asp"
%>
