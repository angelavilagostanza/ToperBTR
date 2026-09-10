<!--#include virtual="/include/bootstrap.asp"-->
<!--#include virtual="/dal/ficheros_dal.asp"-->
<%
Dim conn
Set conn = NuevaConexion()
Call RequiereAccion(conn, CodigoAccion(conn, "LECTURA_FICHEROS"))

Dim esAdministrador, puedeVerIncidencias
esAdministrador     = (Session("CodPerfil") = CodigoPerfil(conn, "Administrador"))
' El perfil Tercero tiene LECTURA_FICHEROS pero no LECTURA_INCIDENCIAS (matriz real
' confirmada, ver docs/04-matriz-permisos.md) - no se enlaza a Incidencias si no puede
' entrar, para no ofrecer un enlace que acabaria en la pagina de "sin permiso".
puedeVerIncidencias = TienePermiso(Session("AccionesPermitidas"), CodigoAccion(conn, "LECTURA_INCIDENCIAS"))

Dim fNombre, fTipo, fOperador, fDesde, fHasta
Dim buscado, limite, pagina, totalRegistros, totalPaginas, offsetReg

fNombre   = Trim(Request.QueryString("nombreFichero"))
fTipo     = Request.QueryString("tipo")
fOperador = Request.QueryString("operador")
fDesde    = Trim(Request.QueryString("fechaDesde"))
fHasta    = Trim(Request.QueryString("fechaHasta"))
buscado   = (Request.QueryString("buscado") = "1")

' Gate servidor: tipo EIR solo para Administrador aunque llegue en el QueryString
If UCase(fTipo & "") = "E" And Not esAdministrador Then fTipo = ""

Select Case CStr(Request.QueryString("limite"))
    Case "100":  limite = 100
    Case "1000": limite = 1000
    Case Else:   limite = 25
End Select

pagina = 1
If IsNumeric(Request.QueryString("pagina")) Then
    If CLng(Request.QueryString("pagina")) > 1 Then
        pagina = CLng(Request.QueryString("pagina"))
    End If
End If

' Mapa clave->nombre de operador para etiquetar filas (el operador se infiere del nombre
' del fichero; el cursor es forward-only, por eso se pide el catalogo dos veces).
Dim rsOperadores, mapaOperadores
Set rsOperadores = ListarOperadores(conn)
Set mapaOperadores = Server.CreateObject("Scripting.Dictionary")
Do While Not rsOperadores.EOF
    If Not mapaOperadores.Exists(rsOperadores("CLAVE") & "") Then
        mapaOperadores.Add rsOperadores("CLAVE") & "", rsOperadores("NOMBRE") & ""
    End If
    rsOperadores.MoveNext
Loop
rsOperadores.Close
Set rsOperadores = ListarOperadores(conn)

Dim rsFicheros
If buscado Then
    totalRegistros = ContarFicheros(conn, fNombre, fTipo, fOperador, fDesde, fHasta, esAdministrador)
    totalPaginas   = Int((totalRegistros + limite - 1) / limite)
    If totalPaginas < 1 Then totalPaginas = 1
    If pagina > totalPaginas Then pagina = totalPaginas
    offsetReg      = (pagina - 1) * limite
    Set rsFicheros = BuscarFicheros(conn, fNombre, fTipo, fOperador, fDesde, fHasta, esAdministrador, limite, offsetReg)
End If

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

Function UrlPagina(nuevaPagina, nuevoLimite)
    Dim qs
    qs = "buscado=1"
    If EsCadenaNoVacia(fNombre)   Then qs = qs & "&nombreFichero=" & Server.URLEncode(fNombre)
    If EsCadenaNoVacia(fTipo)     Then qs = qs & "&tipo="          & Server.URLEncode(fTipo)
    If EsCadenaNoVacia(fOperador) And fOperador <> "-1" Then qs = qs & "&operador=" & Server.URLEncode(fOperador)
    If EsCadenaNoVacia(fDesde)    Then qs = qs & "&fechaDesde="    & Server.URLEncode(fDesde)
    If EsCadenaNoVacia(fHasta)    Then qs = qs & "&fechaHasta="    & Server.URLEncode(fHasta)
    qs = qs & "&limite=" & nuevoLimite & "&pagina=" & nuevaPagina
    UrlPagina = "/modules/ficheros/buscar.asp?" & qs
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

    <form method="get" action="/modules/ficheros/buscar.asp" class="formulario-busqueda">
        <input type="hidden" name="buscado" value="1">
        <div class="formulario-fila">
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
            <label>Registros
                <select name="limite">
                    <option value="25"<% If limite = 25 Then Response.Write " selected" %>>25</option>
                    <option value="100"<% If limite = 100 Then Response.Write " selected" %>>100</option>
                    <option value="1000"<% If limite = 1000 Then Response.Write " selected" %>>1000</option>
                </select>
            </label>
        </div>
        <div class="formulario-fila">
            <label>Creacion desde <input type="date" name="fechaDesde" value="<%= Server.HTMLEncode(fDesde) %>"></label>
            <label>Creacion hasta <input type="date" name="fechaHasta" value="<%= Server.HTMLEncode(fHasta) %>"></label>
        </div>
        <div class="formulario-fila">
            <button type="submit">Filtrar</button>
        </div>
    </form>

    <% If buscado Then %>

    <div class="paginacion">
        <span class="paginacion-info">
            <%= totalRegistros %> resultado<% If totalRegistros <> 1 Then Response.Write "s" %>
            &mdash; P&aacute;gina <%= pagina %> de <%= totalPaginas %>
        </span>
        <nav class="paginacion-nav" aria-label="Paginaci&oacute;n">
            <% If pagina > 1 Then %>
            <a href="<%= UrlPagina(pagina - 1, limite) %>" class="paginacion-link">&larr; Anterior</a>
            <% Else %>
            <span class="paginacion-link deshabilitado">&larr; Anterior</span>
            <% End If %>
            <% If pagina < totalPaginas Then %>
            <a href="<%= UrlPagina(pagina + 1, limite) %>" class="paginacion-link">Siguiente &rarr;</a>
            <% Else %>
            <span class="paginacion-link deshabilitado">Siguiente &rarr;</span>
            <% End If %>
        </nav>
    </div>

    <table class="tabla-datos">
    <thead>
    <tr><th>Nombre</th><th>Operador</th><th>Tipo</th><th>Fecha de creacion</th><th>Incidencias</th></tr>
    </thead>
    <tbody>
    <% If rsFicheros.EOF Then %>
    <tr><td colspan="5" style="text-align:center;color:#777;padding:1.5rem 0">Sin resultados para los filtros indicados.</td></tr>
    <% Else %>
    <% Do While Not rsFicheros.EOF
        Dim nombreFila, nombreFilaJs
        nombreFila   = rsFicheros("NOMBRE") & ""
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
    <% End If %>
    </tbody>
    </table>

    <% End If %>
</main>
</div>
<!--#include virtual="/include/layout/footer.asp"-->
</div>
<%
' formVerIncidencias usa GET para que incidencias/buscar.asp reciba los params
' en QueryString (ese modulo usa Request.QueryString desde que se añadio paginacion).
%>
<form name="formVerIncidencias" method="get" action="/modules/incidencias/buscar.asp" style="display:none">
    <input type="hidden" name="buscado" value="1">
    <input type="hidden" id="nombreFicheroInc" name="nombreFichero" value="">
</form>
</body>
</html>
<%
If buscado Then rsFicheros.Close
rsOperadores.Close
conn.Close
Set conn = Nothing
%>
