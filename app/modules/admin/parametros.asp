<!--#include virtual="/include/bootstrap.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequiereAccion(conn, CodigoAccion(conn, "ADMINISTRADOR"))

' Un token unico para TODOS los formularios de esta pagina.
' GenerarTokenCsrf() sobreescribe el valor en sesion — llamarlo de nuevo por cada fila
' invalidaria los tokens anteriores antes de que el admin pudiera guardarlos.
Dim tokenCsrf
tokenCsrf = GenerarTokenCsrf()

Dim rsParams
Set rsParams = EjecutarConsulta(conn, "SELECT CLAVE, VALOR FROM PARAMETROS_REF ORDER BY CLAVE", Array())

' Descripciones legibles de cada parametro conocido por ToperBTR.
' Los que no aparecen aqui muestran solo su clave, sin descripcion adicional.
Function DescripcionParametro(clave)
    Select Case clave
        Case "COD_FICHERO_GENERICO"      : DescripcionParametro = "Codigo del fichero generico (no editar — rompe la generacion de ficheros)"
        Case "COD_OPER_HOME"             : DescripcionParametro = "Codigo del operador home (Grupo MASMOVIL)"
        Case "DIAS_MARGEN"               : DescripcionParametro = "Dias de margen para validar alta reciente en LISTA_BLANCA"
        Case "DIRECTORIO_DWH"            : DescripcionParametro = "Ruta al directorio de ficheros DWH en el servidor (ruta de servidor — no editar en produccion)"
        Case "DIRECTORIO_EIR"            : DescripcionParametro = "Ruta al directorio de ficheros EIR en el servidor (ruta de servidor — no editar en produccion)"
        Case "DURACION_PASSWORD"         : DescripcionParametro = "Dias de validez de la contrasena de usuario (0 = sin caducidad)"
        Case "EXTENSION_FICHERO_EIR"     : DescripcionParametro = "Extension de los ficheros EIR generados por EIRControl"
        Case "NUM_FALLOS_LOGIN_USUARIO"  : DescripcionParametro = "Intentos de login fallidos — parametro legacy, actualmente inoperante en ToperBTR"
        Case "NUM_MAX_INTENTOS_FALLIDOS" : DescripcionParametro = "Intentos de login fallidos antes de bloquear la cuenta de usuario"
        Case "NUM_OLD_PASSWORD"          : DescripcionParametro = "Numero de contrasenas anteriores que no se pueden reutilizar"
        Case "RAIZ_FICHEROS"             : DescripcionParametro = "Ruta raiz del directorio de ficheros (ruta de servidor — no editar en produccion)"
        Case Else                        : DescripcionParametro = ""
    End Select
End Function
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Administracion</title>
<link rel="stylesheet" href="/assets/css/site.css">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Administracion — Parametros del sistema</h2>

    <p class="aviso-info">
        Modifique los valores con precaucion. Los parametros de rutas de servidor
        (<code>RAIZ_FICHEROS</code>, <code>DIRECTORIO_EIR</code>, <code>DIRECTORIO_DWH</code>)
        afectan directamente a EIRControl y no deben cambiarse en produccion sin coordinacion previa.
    </p>

    <table class="tabla-datos tabla-admin-params">
    <thead>
    <tr>
        <th>Clave</th>
        <th>Valor actual</th>
        <th>Descripcion</th>
        <th></th>
    </tr>
    </thead>
    <tbody>
    <%
    Do While Not rsParams.EOF
        Dim clave, valor, desc
        clave = rsParams("CLAVE") & ""
        valor = rsParams("VALOR") & ""
        desc  = DescripcionParametro(clave)
    %>
    <tr>
        <td><code><%= Server.HTMLEncode(clave) %></code></td>
        <td>
            <form method="post" action="/modules/admin/parametro_do.asp" class="form-param-inline">
                <input type="hidden" name="csrf"  value="<%= Server.HTMLEncode(tokenCsrf) %>">
                <input type="hidden" name="clave" value="<%= Server.HTMLEncode(clave) %>">
                <input type="text"   name="valor" value="<%= Server.HTMLEncode(valor) %>"
                       maxlength="500" class="input-param-valor">
                <button type="submit" class="btn btn-secundario btn-sm">Guardar</button>
            </form>
        </td>
        <td class="desc-param"><%= Server.HTMLEncode(desc) %></td>
    </tr>
    <%
        rsParams.MoveNext
    Loop
    %>
    </tbody>
    </table>

</main>
</div>
<!--#include virtual="/include/layout/footer.asp"-->
</div>
</body>
</html>
<%
rsParams.Close
conn.Close
Set conn = Nothing
%>
