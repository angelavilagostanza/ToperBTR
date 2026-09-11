<!--#include virtual="/include/bootstrap.asp"-->
<%
Call RequiereSesion()
%>
<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<title>ToperBTR - Inicio</title>
<link rel="stylesheet" href="/assets/css/site.css">
<link rel="icon" href="/images/ToperBTR_favicon.ico" type="image/x-icon">
</head>
<body>
<div class="app-shell">
<!--#include virtual="/include/layout/header.asp"-->
<div class="app-body">
<!--#include virtual="/include/layout/menu.asp"-->
<main class="contenido">
    <div class="home-intro">
        <h2>Bienvenido, <%= Server.HTMLEncode(Session("Nombre")) %></h2>
        <p>BTR es una aplicación para gestionar el Bloqueo de Terminales Robados.</p>
    </div>

    <div class="indice-funcionalidades">

        <% If puedeSolicitudes Then %>
        <a href="/modules/solicitudes/consulta.asp" class="indice-card">
            <svg class="indice-card-icono" xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <path d="M16 4h2a2 2 0 012 2v14a2 2 0 01-2 2H6a2 2 0 01-2-2V6a2 2 0 012-2h2"/><rect x="8" y="2" width="8" height="4" rx="1"/>
            </svg>
            <span class="indice-card-cuerpo">
                <span class="indice-card-titulo">Solicitudes</span>
                <span class="indice-card-desc">Consulta y gestión de solicitudes de Inclusión y Exclusión de terminales en la base.</span>
            </span>
        </a>
        <% End If %>

        <% If puedeIncidencias Then %>
        <a href="/modules/incidencias/buscar.asp" class="indice-card">
            <svg class="indice-card-icono" xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
            </svg>
            <span class="indice-card-cuerpo">
                <span class="indice-card-titulo">Incidencias</span>
                <span class="indice-card-desc">Consulta de incidencias sobre terminales. Visualización de detalle y comentarios por incidencia.</span>
            </span>
        </a>
        <% End If %>

        <% If puedeFicheros Then %>
        <a href="/modules/ficheros/buscar.asp" class="indice-card">
            <svg class="indice-card-icono" xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <path d="M22 19a2 2 0 01-2 2H4a2 2 0 01-2-2V5a2 2 0 012-2h5l2 3h9a2 2 0 012 2z"/>
            </svg>
            <span class="indice-card-cuerpo">
                <span class="indice-card-titulo">Ficheros</span>
                <span class="indice-card-desc">Búsqueda y descarga de ficheros de datos intercambiados entre operadores.</span>
            </span>
        </a>
        <% End If %>

        <% If puedeBloqueos Then %>
        <a href="/modules/bloqueos/bloqueo.asp" class="indice-card">
            <svg class="indice-card-icono" xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0110 0v4"/>
            </svg>
            <span class="indice-card-cuerpo">
                <span class="indice-card-titulo">Bloqueos</span>
                <span class="indice-card-desc">Bloqueo y desbloqueo directo de terminales. Consulta del estado de bloqueo de un terminal.</span>
            </span>
        </a>
        <% End If %>

        <% If puedeUsuarios Then %>
        <a href="/modules/usuarios/listar.asp" class="indice-card">
            <svg class="indice-card-icono" xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/>
            </svg>
            <span class="indice-card-cuerpo">
                <span class="indice-card-titulo">Usuarios</span>
                <span class="indice-card-desc">Alta, baja y desbloqueo de cuentas de usuario. Cambio de contraseña por el administrador.</span>
            </span>
        </a>
        <% End If %>

        <% If puedeAdmin Then %>
        <a href="/modules/admin/parametros.asp" class="indice-card">
            <svg class="indice-card-icono" xmlns="http://www.w3.org/2000/svg" width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 012.83-2.83l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 010 4h-.09a1.65 1.65 0 00-1.51 1z"/>
            </svg>
            <span class="indice-card-cuerpo">
                <span class="indice-card-titulo">Parámetros</span>
                <span class="indice-card-desc">Configuración y edición de los parámetros de referencia del sistema BTR.</span>
            </span>
        </a>
        <% End If %>

    </div>
</main>
</div>
<!--#include virtual="/include/layout/footer.asp"-->
</div>
</body>
</html>
