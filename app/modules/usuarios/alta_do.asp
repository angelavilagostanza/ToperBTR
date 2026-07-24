<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/usuarios_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequierePerfil(Array(CodigoPerfil(conn, "Seguridad")))

If Not TokenCsrfValido(Request.Form("csrf")) Then
    conn.Close
    Response.Redirect "/errors/error.asp?msg=csrf"
    Response.End
End If

Dim usuario, nombre, apellidos, password, password2, codPerfil
usuario = Trim(Request.Form("usuario"))
nombre = Trim(Request.Form("nombre"))
apellidos = Trim(Request.Form("apellidos"))
password = Request.Form("password")
password2 = Request.Form("password2")
codPerfil = Request.Form("perfil")

Dim errorMsg
errorMsg = ""

If Not (EsCadenaNoVacia(usuario) And LongitudEntre(usuario, 1, 15)) Then
    errorMsg = "El login es obligatorio (maximo 15 caracteres)."
ElseIf Not EsCadenaNoVacia(nombre) Or Not EsCadenaNoVacia(apellidos) Then
    errorMsg = "Nombre y apellidos son obligatorios."
ElseIf Not IsNumeric(codPerfil) Then
    errorMsg = "Debe seleccionar un perfil."
ElseIf Len(password) < 10 Then
    errorMsg = "La contrasena debe tener al menos 10 caracteres."
ElseIf password <> password2 Then
    errorMsg = "Las dos contrasenas no coinciden."
ElseIf ExisteLogin(conn, usuario) Then
    ' Decision de negocio: un login dado de baja NO se puede reutilizar (por trazabilidad),
    ' a diferencia del legacy, que si lo permitia. Por eso aqui NO se excluye ESTADO=Borrado.
    errorMsg = "Ese login ya existe (activo o dado de baja) y no se puede reutilizar."
End If

If errorMsg <> "" Then
    conn.Close
    Set conn = Nothing
    Response.Redirect "/modules/usuarios/alta.asp?error=" & Server.URLEncode(errorMsg)
    Response.End
End If

Dim codOperador, maximoUsuarios, actualesUsuarios, codBorrado
codOperador = Session("CodOperador")
codBorrado = CodigoEstadoUsuario(conn, "Borrado")
actualesUsuarios = ContarUsuariosActivosOperador(conn, codOperador, codBorrado)
maximoUsuarios = NumMaxUsuariosOperador(conn, codOperador)

If actualesUsuarios >= maximoUsuarios Then
    conn.Close
    Set conn = Nothing
    Response.Redirect "/modules/usuarios/alta.asp?error=" & Server.URLEncode("Se ha alcanzado el numero maximo de usuarios para su operador.")
    Response.End
End If

Dim hashPassword, fechaCaducidad, diasCaducidad, codActivo
diasCaducidad = ValorParametroEntero(conn, "DURACION_PASSWORD", 60)
fechaCaducidad = DateAdd("d", diasCaducidad, Now())
hashPassword = CrearHashPbkdf2(password, CLng(Application("PBKDF2Iteraciones")))
codActivo = CodigoEstadoUsuario(conn, "Activo")

Call CrearUsuario(conn, usuario, nombre, apellidos, hashPassword, "PBKDF2SHA256", fechaCaducidad, CLng(codPerfil), codOperador, codActivo)

Dim nuevoId
nuevoId = ObtenerUltimoId(conn)

Call RegistrarAuditoria(conn, Session("CodIndiceUsuario"), "USUARIO_ALTA", "USUARIOS", nuevoId, "OK", "Login creado: " & usuario)

conn.Close
Set conn = Nothing

Call EstablecerMensajeFlash("exito", "Usuario creado correctamente.")
Response.Redirect "/modules/usuarios/listar.asp"
%>
