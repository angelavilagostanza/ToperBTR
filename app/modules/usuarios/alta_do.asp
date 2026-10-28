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

' Hallazgo de seguridad M1: antes solo se exigia longitud, sin restringir que caracteres
' podia contener el login - permitia dar de alta un usuario con "<"/">"/comillas en el
' login, que luego se escribia SIN escapar en un comentario HTML de usuarios/listar.asp
' (Response.Write de depuracion, ya eliminado), lo que permitia a un administrador de
' Seguridad inyectar JavaScript que se ejecutaba en el navegador de OTRO administrador de
' Seguridad al abrir ese listado (XSS almacenado). El patron de abajo solo admite los
' caracteres habituales de un login/email (se amplio a 50 caracteres pensando en que el
' login fuera un email - commit 84f413e - pero la validacion de formato email en si sigue
' pendiente como tarea aparte, ver memoria del proyecto); basta para que el login nunca
' pueda contener "<", ">", comillas ni espacios, cerrando el vector de raiz ademas de
' quitar el Response.Write que lo exponia.
If Not (EsCadenaNoVacia(usuario) And LongitudEntre(usuario, 1, 50) And CumpleFormato(usuario, "^[A-Za-z0-9._%+@-]+$")) Then
    errorMsg = "El login es obligatorio (maximo 50 caracteres; solo letras, numeros, y . _ % + @ -)."
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
Response.End
%>
