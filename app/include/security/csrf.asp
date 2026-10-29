<%
Function GenerarTokenCsrf()
    Dim token
    token = GenerarTokenAleatorio(32)
    Session("CsrfToken") = token
    GenerarTokenCsrf = token
End Function

' Hallazgo B3: antes comparaba con "=" normal, que se detiene en el primer caracter
' distinto - en teoria permite un timing attack (medir el tiempo de respuesta para
' deducir el token caracter a caracter en vez de tener que adivinarlo entero). Riesgo
' bajo en la practica (hace falta medir diferencias de microsegundos de forma muy
' fiable por red), pero ya existe ComparacionSegura (crypto_pbkdf2.asp) para este mismo
' proposito con los hashes de password - se reutiliza aqui por consistencia, sin coste.
Function TokenCsrfValido(tokenRecibido)
    TokenCsrfValido = False
    If Not IsEmpty(Session("CsrfToken")) Then
        If Session("CsrfToken") <> "" And ComparacionSegura(tokenRecibido & "", Session("CsrfToken") & "") Then
            TokenCsrfValido = True
        End If
    End If
End Function
%>
