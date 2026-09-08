<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/ficheros_dal.asp"-->
<%
' Corrige el hallazgo de seguridad mas grave confirmado en el codigo Java real
' (ActionFicheros.java): el chequeo de sesion estaba comentado por completo y el fichero
' ya se habia volcado a la respuesta antes de llegar a la (inexistente) comprobacion -
' cualquiera con la URL/nombre de fichero podia descargarlo sin login. Aqui el gate es
' real y se ejecuta ANTES de tocar el disco: RequiereAccion corta la peticion si no hay
' sesion valida o si el perfil no tiene LECTURA_FICHEROS.
Dim conn
Set conn = NuevaConexion()
Call RequiereAccion(conn, CodigoAccion(conn, "LECTURA_FICHEROS"))

Dim nombreSolicitado, fichero
nombreSolicitado = Trim(Request.QueryString("nombre"))

If Not EsCadenaNoVacia(nombreSolicitado) Then
    conn.Close
    Set conn = Nothing
    Response.Redirect "/errors/error.asp?msg=ficherosindato"
    Response.End
End If

' La ruta fisica se construye siempre como Server.MapPath del directorio virtual IIS
' "ficherosBTR" mas el nombre+extension que devuelve la BD - nunca se concatena input
' del usuario, evitando cualquier traversal de rutas.
Set fichero = ObtenerFicheroPorNombre(conn, nombreSolicitado)

If fichero Is Nothing Then
    Call RegistrarAuditoria(conn, Session("CodIndiceUsuario"), "FICHERO_DESCARGA", "FICHEROS", Null, "NO_ENCONTRADO", "Nombre solicitado: " & nombreSolicitado)
    conn.Close
    Set conn = Nothing
    Response.Redirect "/errors/error.asp?msg=ficheronoencontrado"
    Response.End
End If

Dim rutaCompleta, nombreDescarga
nombreDescarga = fichero("NombreFichero")
If EsCadenaNoVacia(fichero("Extension")) Then nombreDescarga = nombreDescarga & "." & fichero("Extension")
rutaCompleta = Server.MapPath("/ficherosBTR/" & nombreDescarga)

Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(rutaCompleta) Then
    Call RegistrarAuditoria(conn, Session("CodIndiceUsuario"), "FICHERO_DESCARGA", "FICHEROS", fichero("CodFichero"), "FICHERO_FISICO_NO_ENCONTRADO", rutaCompleta)
    Set fso = Nothing
    conn.Close
    Set conn = Nothing
    Response.Redirect "/errors/error.asp?msg=ficheronoencontrado"
    Response.End
End If
Set fso = Nothing

Call RegistrarAuditoria(conn, Session("CodIndiceUsuario"), "FICHERO_DESCARGA", "FICHEROS", fichero("CodFichero"), "OK", rutaCompleta)
conn.Close
Set conn = Nothing

' Quitamos CR/LF por si acaso, para no arrastrar una inyeccion de cabeceras HTTP aunque
' el valor venga de BD y no directamente del usuario (defensa en profundidad).
nombreDescarga = Replace(Replace(nombreDescarga, Chr(13), ""), Chr(10), "")

Dim stream
Set stream = Server.CreateObject("ADODB.Stream")
stream.Type = 1 ' adTypeBinary
stream.Open
stream.LoadFromFile rutaCompleta

Response.Clear
Response.ContentType = "application/octet-stream"
Response.AddHeader "Content-Disposition", "attachment; filename=" & nombreDescarga
Response.BinaryWrite stream.Read()
stream.Close
Set stream = Nothing
Response.End
%>
