<%
Function GenerarTokenCsrf()
    Dim token
    token = GenerarTokenAleatorio(32)
    Session("CsrfToken") = token
    GenerarTokenCsrf = token
End Function

Function TokenCsrfValido(tokenRecibido)
    TokenCsrfValido = False
    If Not IsEmpty(Session("CsrfToken")) Then
        If Session("CsrfToken") <> "" And tokenRecibido = Session("CsrfToken") Then
            TokenCsrfValido = True
        End If
    End If
End Function
%>
