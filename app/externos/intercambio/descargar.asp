<%@ Language="VBScript" EnableSessionState="False" %>
<%
Option Explicit
' Endpoint maquina-a-maquina para que los sistemas de otros operadores descarguen el
' fichero diario de intercambio EIR (altas/bajas de IMEI en lista negra) por nombre
' exacto. Reemplaza a ficheros.do?cmd=publicarFichero&action=abrir&nombreFichero=X del
' legacy (EIRWeb), que en produccion vivia detras de un frontal que exigia certificado
' de cliente (mTLS) antes de llegar al endpoint interno real - el chequeo de sesion de
' ese endpoint interno estaba comentado a proposito, porque la autenticacion ya habia
' ocurrido en el salto anterior.
'
' Aqui replicamos ese mismo modelo de confianza: IIS debe exigir SSL + certificado de
' cliente en esta ruta ANTES de que la peticion llegue a este fichero (ver web.config de
' esta carpeta y docs/12-endpoint-intercambio-operadores.md para el checklist del
' administrador). Este script NO inspecciona Request.ClientCertificate - confia en que,
' si se esta ejecutando, es porque IIS ya acepto un certificado valido.
'
' No hay sesion de navegador (EnableSessionState=False arriba evita que IIS cree una
' sesion por cada llamada sin motivo), así que tampoco se usan las funciones Requiere*
' (todas dependen de Session). Por el mismo motivo, NUNCA se redirige a
' /errors/error.asp (pagina HTML pensada para un navegador con sesion) - toda respuesta
' de error aqui es un codigo HTTP + texto plano, usando las variantes "Tx" (que no
' redirigen solas) de conexion.asp/ficheros_dal.asp.
%>
<!--#include virtual="/include/data/conexion.asp"-->
<!--#include virtual="/include/data/catalogos.asp"-->
<!--#include virtual="/include/common/validacion.asp"-->
<!--#include virtual="/include/security/auditoria.asp"-->
<!--#include virtual="/dal/ficheros_dal.asp"-->
<%
Sub ResponderTexto(codigoEstado, mensaje)
    Response.Status = codigoEstado
    Response.ContentType = "text/plain"
    Response.Write mensaje
    Response.End
End Sub

If UCase(Request.ServerVariables("REQUEST_METHOD")) <> "GET" Then
    Response.AddHeader "Allow", "GET"
    Call ResponderTexto("405 Method Not Allowed", "Method not allowed")
End If

Dim conn
Set conn = NuevaConexionTx()
If conn Is Nothing Then
    Call ResponderTexto("500 Internal Server Error", "Service unavailable")
End If

Dim nombreCrudo, nombreNormalizado, esValido, letraOperador
nombreCrudo = Trim(Request.QueryString("fichero") & "")
esValido = False

' Formato de los ficheros reales de intercambio diario: 1 letra de operador + 8 digitos
' de fecha (p.ej. Y20260802, A20260729) - excluye a proposito los ficheros "Consolidacion"
' (<9 caracteres) y "EIR" (11 caracteres, feed propio de Yoigo, solo visible para
' Administrador en el buscador interno) que no son los que piden los otros operadores.
If EsCadenaNoVacia(nombreCrudo) Then
    If CumpleFormato(nombreCrudo, "^[A-Za-z]\d{8}$") Then
        letraOperador = UCase(Left(nombreCrudo, 1))
        nombreNormalizado = letraOperador & Mid(nombreCrudo, 2)

        Dim rsOperadores, clavesValidas
        Set clavesValidas = Server.CreateObject("Scripting.Dictionary")
        Set rsOperadores = ListarOperadores(conn)
        Do While Not rsOperadores.EOF
            clavesValidas(UCase(rsOperadores("CLAVE") & "")) = True
            rsOperadores.MoveNext
        Loop
        rsOperadores.Close

        If clavesValidas.Exists(letraOperador) Then esValido = True
    End If
End If

If Not esValido Then
    Call RegistrarAuditoria(conn, Null, "FICHERO_DESCARGA_INTERCAMBIO", "FICHEROS", Null, _
        "RECHAZADO_FORMATO", "Nombre solicitado: " & nombreCrudo)
    conn.Close
    Set conn = Nothing
    Call ResponderTexto("400 Bad Request", "Invalid filename")
End If

Dim fichero
Set fichero = ObtenerFicheroPorNombreTx(conn, nombreNormalizado)

If fichero Is Nothing Then
    Call RegistrarAuditoria(conn, Null, "FICHERO_DESCARGA_INTERCAMBIO", "FICHEROS", Null, _
        "NO_ENCONTRADO", "Nombre solicitado: " & nombreCrudo)
    conn.Close
    Set conn = Nothing
    Call ResponderTexto("404 Not Found", "File not found")
End If

' Misma resolucion de ruta fisica que ficheros/descargar.asp: siempre a partir del
' nombre+extension que devuelve la BD sobre el virtual directory IIS "ficherosBTR",
' nunca concatenando input del usuario - evita cualquier traversal de rutas.
Dim rutaCompleta, nombreDescarga
nombreDescarga = fichero("NombreFichero")
If EsCadenaNoVacia(fichero("Extension")) Then nombreDescarga = nombreDescarga & "." & fichero("Extension")
rutaCompleta = Server.MapPath("/ficherosBTR/" & nombreDescarga)

Dim fso
Set fso = Server.CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(rutaCompleta) Then
    Call RegistrarAuditoria(conn, Null, "FICHERO_DESCARGA_INTERCAMBIO", "FICHEROS", fichero("CodFichero"), _
        "NO_ENCONTRADO_FISICO", rutaCompleta)
    Set fso = Nothing
    conn.Close
    Set conn = Nothing
    Call ResponderTexto("404 Not Found", "File not found")
End If
Set fso = Nothing

Call RegistrarAuditoria(conn, Null, "FICHERO_DESCARGA_INTERCAMBIO", "FICHEROS", fichero("CodFichero"), "OK", rutaCompleta)
conn.Close
Set conn = Nothing

' Quitamos CR/LF por si acaso, para no arrastrar una inyeccion de cabeceras HTTP aunque
' el valor venga de BD y ya haya pasado la validacion de formato (defensa en profundidad,
' igual que en ficheros/descargar.asp).
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
