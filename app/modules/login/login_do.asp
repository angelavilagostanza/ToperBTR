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
' de ASP clasico/IIS; mitigada con timeout corto y con la cookie ToperBTR_Fp de abajo.
Session("Autenticado") = True
Session("CodIndiceUsuario") = resultado("CodIndiceUsuario")
Session("CodUsuario") = resultado("CodUsuario")
Session("Nombre") = resultado("Nombre")
Session("Apellidos") = resultado("Apellidos")
Session("CodPerfil") = resultado("CodPerfil")
Session("NombrePerfil") = resultado("NombrePerfil")
Session("CodOperador") = resultado("CodOperador")
Session("AccionesPermitidas") = resultado("AccionesPermitidas")

' ASP clasico no permite marcar la cookie automatica ASPSESSIONID... como HttpOnly (esa
' opcion no existe en IIS para ASP clasico, solo para ASP.NET), asi que esa cookie sigue
' siendo legible por JavaScript. Para que un robo de ASPSESSIONID via XSS no sirva para
' suplantar la sesion, se exige ademas esta segunda cookie SI HttpOnly, con un valor
' aleatorio que se valida en cada peticion (RequiereSesion, autorizacion.asp) contra
' Session("Fingerprint") - un atacante que solo robe el ASPSESSIONID nunca podra leer
' (ni por tanto reenviar) esta cookie.
Session("Fingerprint") = GenerarTokenAleatorio(32)
Response.Cookies("ToperBTR_Fp") = Session("Fingerprint")
Response.Cookies("ToperBTR_Fp").HttpOnly = True
If UCase(Request.ServerVariables("HTTPS")) = "ON" Then
    Response.Cookies("ToperBTR_Fp").Secure = True
End If

Response.Redirect "/modules/home/home.asp"
Response.End
%>
