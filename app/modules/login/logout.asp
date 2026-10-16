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
Response.Redirect "/modules/login/login.asp"
%>
