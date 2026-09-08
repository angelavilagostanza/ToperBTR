<!--#include virtual="/include/bootstrap.asp"-->
<%
' Timeout extendido para permitir que los hashes legacy (50000 iter) puedan verificarse
' y migrarse en el primer login. Eliminar una vez que todos los usuarios hayan hecho
' su primer login con el nuevo servidor (o cuando los hashes DB se restablezcan a 200 iter).
Server.ScriptTimeout = 360

Dim csrfRecibido, codUsuario, password
csrfRecibido = Request.Form("csrf")
codUsuario = Trim(Request.Form("usuario"))
password = Request.Form("password")

If Not TokenCsrfValido(csrfRecibido) Then
    Response.Redirect "/modules/login/login.asp?error=" & Server.URLEncode("La sesion del formulario ha expirado, intentelo de nuevo.")
    Response.End
End If

If Not EsCadenaNoVacia(codUsuario) Or Not EsCadenaNoVacia(password) Then
    Response.Redirect "/modules/login/login.asp?error=" & Server.URLEncode("Usuario o contrasena incorrectos.")
    Response.End
End If

Dim conn
Set conn = NuevaConexion()

Dim resultado
Set resultado = IntentarLogin(conn, codUsuario, password)

conn.Close
Set conn = Nothing

If Not resultado("Exito") Then
    Response.Redirect "/modules/login/login.asp?error=" & Server.URLEncode(resultado("Mensaje"))
    Response.End
End If

' Nota: no se usa Session.Abandon antes de reescribir la sesion (rotacion de SessionID
' tras login) porque en ASP clasico Session.Abandon destruye el estado al FINAL de esta
' misma peticion, incluido lo que se escriba despues en el mismo request - se perderia
' todo lo que fijamos a continuacion. Limitacion conocida y aceptada del modelo de sesion
' de ASP clasico/IIS; mitigada con timeout corto y cookie de sesion HttpOnly/Secure.
Session("Autenticado") = True
Session("CodIndiceUsuario") = resultado("CodIndiceUsuario")
Session("CodUsuario") = resultado("CodUsuario")
Session("Nombre") = resultado("Nombre")
Session("Apellidos") = resultado("Apellidos")
Session("CodPerfil") = resultado("CodPerfil")
Session("CodOperador") = resultado("CodOperador")
Session("AccionesPermitidas") = resultado("AccionesPermitidas")

Response.Redirect "/modules/home/home.asp"
Response.End
%>
