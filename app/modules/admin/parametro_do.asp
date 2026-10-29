<!--#include virtual="/include/bootstrap.asp"-->
<%
Response.Buffer = True

Dim conn
Set conn = NuevaConexion()
Call RequiereAccion(conn, CodigoAccion(conn, "ADMINISTRADOR"))

If Not TokenCsrfValido(Request.Form("csrf")) Then
    Response.Redirect "/errors/error.asp?msg=csrf"
    Response.End
End If

Dim clave, nuevoValor
clave      = Trim(Request.Form("clave") & "")
nuevoValor = Trim(Request.Form("valor") & "")

' La clave no puede llegar vacia: seria un fallo del formulario o manipulacion.
If Not EsCadenaNoVacia(clave) Then
    Call EstablecerMensajeFlash("error", "No se ha indicado el parametro a modificar.")
    Response.Redirect "/modules/admin/parametros.asp"
    Response.End
End If

' Verificar que la clave existe y obtener el valor anterior para auditoria.
' Si el SELECT no devuelve filas, la clave no existe — no se permite crear
' parametros nuevos (podria romper la compatibilidad con EIRControl).
Dim rsExiste, valorAnterior
Set rsExiste = EjecutarConsulta(conn, _
    "SELECT VALOR FROM PARAMETROS_REF WHERE CLAVE = ?", Array(clave))

If rsExiste.EOF Then
    rsExiste.Close
    Call EstablecerMensajeFlash("error", "El parametro '" & clave & "' no existe. No se pueden crear parametros nuevos desde esta pantalla.")
    Response.Redirect "/modules/admin/parametros.asp"
    Response.End
End If

valorAnterior = rsExiste("VALOR") & ""
rsExiste.Close

' Hallazgo B5: las claves que la app lee via ValorParametroEntero ya son defensivas (si
' el valor guardado no es numerico, cae en silencio al valor por defecto - no rompe
' nada), pero eso mismo es el problema: un valor invalido se guarda sin aviso y el
' administrador puede pensar que su cambio surtio efecto cuando en realidad fue
' ignorado. Se valida aqui, antes de guardar, para las claves que ToperBTR sabe que
' espera un entero - el resto de claves (incluidas las de ruta de servidor, que ni
' siquiera lee ToperBTR, ver parametros.asp) se dejan como texto libre igual que antes.
Dim clavesEnteras
clavesEnteras = "|COD_FICHERO_GENERICO|NUM_MAX_INTENTOS_FALLIDOS|DURACION_PASSWORD|NUM_OLD_PASSWORD|"
If InStr(1, clavesEnteras, "|" & clave & "|", 1) > 0 Then
    If Not IsNumeric(nuevoValor) Then
        conn.Close
        Set conn = Nothing
        Call EstablecerMensajeFlash("error", "El parametro '" & clave & "' debe ser un numero entero.")
        Response.Redirect "/modules/admin/parametros.asp"
        Response.End
    End If
End If

' UPDATE unico, sin transaccion (sentencia atomica sobre una sola fila).
Call EjecutarNonQuery(conn, _
    "UPDATE PARAMETROS_REF SET VALOR = ? WHERE CLAVE = ?", _
    Array(nuevoValor, clave))

Call RegistrarAuditoria(conn, Session("CodIndiceUsuario"), _
    "MODIFICAR", "PARAMETROS_REF", clave, "OK", _
    "ANTERIOR=" & valorAnterior & " NUEVO=" & nuevoValor)

conn.Close
Set conn = Nothing

Call EstablecerMensajeFlash("ok", "Parametro '" & clave & "' actualizado correctamente.")
Response.Redirect "/modules/admin/parametros.asp"
Response.End
%>
