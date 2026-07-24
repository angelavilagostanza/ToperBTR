<!--#include virtual="/include/bootstrap.asp"-->
<%
If Session("Autenticado") = True Then
    Response.Redirect "/modules/home/home.asp"
Else
    Response.Redirect "/modules/login/login.asp"
End If
%>
