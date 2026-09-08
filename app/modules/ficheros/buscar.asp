<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/ficheros_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequiereAccion(conn, CodigoAccion(conn, "LECTURA_FICHEROS"))

Dim esAdministrador, puedeVerIncidencias
esAdministrador = (Session("CodPerfil") = CodigoPerfil(conn, "Administrador"))
' El perfil Tercero tiene LECTURA_FICHEROS pero no LECTURA_INCIDENCIAS (matriz real
' confirmada, ver docs/04-matriz-permisos.md) - no se enlaza a Incidencias si no puede
' entrar, para no ofrecer un enlace que acabaria en la pagina de "sin permiso".
puedeVerIncidencias = TienePermiso(Session("AccionesPermitidas"), CodigoAccion(conn, "LECTURA_INCIDENCIAS"))

Dim fNombre, fTipo, fOperador, fDesde, fHasta
fNombre = Trim(Request.Form("nombreFichero"))
fTipo = Request.Form("tipo")
fOperador = Request.Form("operador")
fDesde = Trim(Request.Form("fechaDesde"))
fHasta = Trim(Request.Form("fechaHasta"))

' Ficheros EIR (solo Administrador) no pueden filtrarse si el formulario los oculta,
' pero por si acaso llega tipo=E en el POST de alguien sin perfil Administrador, se
' ignora aqui en vez de confiar en que el formulario nunca lo hubiera enviado.
If UCase(fTipo & "") = "E" And Not esAdministrador Then fTipo = ""

Dim rsFicheros, rsOperadores
Set rsOperadores = ListarOperadores(conn)

' Diccionario clave->nombre de operador para etiquetar cada fila: el operador propietario
' de un fichero no es una columna, se infiere del primer caracter del nombre (igual que
' hacia el codigo legacy real). El cursor de rsOperadores es forward-only (no admite
' MoveFirst), asi que se vuelve a pedir el catalogo mas abajo para pintar el desplegable.
Dim mapaOperadores
Set mapaOperadores = Server.CreateObject("Scripting.Dictionary")
Do While Not rsOperadores.EOF
    If Not mapaOperadores.Exists(rsOperadores("CLAVE") & "") Then
        mapaOperadores.Add rsOperadores("CLAVE") & "", rsOperadores("NOMBRE") & ""
    End If
    rsOperadores.MoveNext
Loop
rsOperadores.Close
Set rsOperadores = ListarOperadores(conn)

Set rsFicheros = BuscarFicheros(conn, fNombre, fTipo, fOperador, fDesde, fHasta, esAdministrador)

Function EtiquetaTipoFichero(nombre)
    If Len(nombre) < 9 Then
        EtiquetaTipoFichero = "Consolidacion"
    ElseIf InStr(1, nombre, "EIR", 1) > 0 Then
        EtiquetaTipoFichero = "EIR"
    Else
        EtiquetaTipoFichero = "Diario"
    End If
End Function

Function EtiquetaOperadorFichero(nombre)
    Dim clave
    clave = Left(nombre, 1)
    If mapaOperadores.Exists(clave) Then
        EtiquetaOperadorFichero = mapaOperadores(clave)
    Else
        EtiquetaOperadorFichero = "Grupo MASMOVIL"
    End If
End Function
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Ficheros</title>
<link rel="stylesheet" href="/assets/css/site.css">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <h2>Ficheros</h2>

    <form method="post" action="/modules/ficheros/buscar.asp" class="formulario-busqueda">
        <label>Nombre de fichero (exacto) <input type="text" name="nombreFichero" maxlength="11" value="<%= Server.HTMLEncode(fNombre) %>"></label>
        <label>Tipo
            <select name="tipo">
                <option value="">Todos</option>
                <option value="D"<% If UCase(fTipo & "") = "D" Then Response.Write " selected" %>>Diarios</option>
                <option value="C"<% If UCase(fTipo & "") = "C" Then Response.Write " selected" %>>Consolidacion</option>
                <% If esAdministrador Then %>
                <option value="E"<% If UCase(fTipo & "") = "E" Then Response.Write " selected" %>>Ficheros EIR</option>
                <% End If %>
            </select>
        </label>
        <label>Operador
            <select name="operador">
                <option value="-1">Todos los operadores</option>
                <% Do While Not rsOperadores.EOF %>
                <option value="<%= rsOperadores("COD_OPERADOR") %>"<% If CStr(fOperador) = CStr(rsOperadores("COD_OPERADOR")) Then Response.Write " selected" %>><%= Server.HTMLEncode(rsOperadores("NOMBRE")) %></option>
                <% rsOperadores.MoveNext
                Loop %>
            </select>
        </label>
        <label>Creacion desde <input type="date" name="fechaDesde" value="<%= Server.HTMLEncode(fDesde) %>"></label>
        <label>Creacion hasta <input type="date" name="fechaHasta" value="<%= Server.HTMLEncode(fHasta) %>"></label>
        <button type="submit">Filtrar</button>
    </form>

    <table class="tabla-datos">
    <thead>
    <tr><th>Nombre</th><th>Operador</th><th>Tipo</th><th>Fecha de creacion</th><th>Incidencias</th></tr>
    </thead>
    <tbody>
    <% Do While Not rsFicheros.EOF
        Dim nombreFila, nombreFilaJs
        nombreFila = rsFicheros("NOMBRE") & ""
        nombreFilaJs = Replace(Replace(nombreFila, "\", ""), "'", "")
    %>
    <tr>
        <td><a href="/modules/ficheros/descargar.asp?nombre=<%= Server.URLEncode(nombreFila) %>"><%= Server.HTMLEncode(nombreFila) %></a></td>
        <td><%= Server.HTMLEncode(EtiquetaOperadorFichero(nombreFila)) %></td>
        <td><%= EtiquetaTipoFichero(nombreFila) %></td>
        <td><%= rsFicheros("FECHA_CREACION") %></td>
        <td>
            <% If CLng(rsFicheros("NUM_INCIDENCIAS")) = 0 Then %>
            Sin incidencias
            <% ElseIf puedeVerIncidencias Then %>
            <a href="#" onclick="document.getElementById('nombreFicheroInc').value='<%= nombreFilaJs %>'; document.formVerIncidencias.submit(); return false;">Incidencias: <%= rsFicheros("NUM_INCIDENCIAS") %></a>
            <% Else %>
            Incidencias: <%= rsFicheros("NUM_INCIDENCIAS") %>
            <% End If %>
        </td>
    </tr>
    <% rsFicheros.MoveNext
    Loop %>
    </tbody>
    </table>
</main>
</div>
<!--#include virtual="/include/layout/footer.asp"-->
</div>
<form name="formVerIncidencias" method="post" action="/modules/incidencias/buscar.asp" style="display:none">
    <input type="hidden" id="nombreFicheroInc" name="nombreFichero" value="">
</form>
</body>
</html>
<%
rsFicheros.Close
rsOperadores.Close
conn.Close
Set conn = Nothing
%>
