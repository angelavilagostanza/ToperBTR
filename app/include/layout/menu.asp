<%
' El menu resuelve los permisos sin consultar la BD: los codigos de catalogo (ACCION_REF,
' PERFIL_REF) se cargan en Application scope al arrancar (global.asa); las acciones
' permitidas del usuario vienen de Session (cargadas en el login). Usuarios y Bloqueos
' no tienen fila en ACCION_PERMITIDA_REF en el sistema real (confirmado contra produccion):
' se gatean aqui explicitamente por perfil, en el UNICO sitio donde vive esa regla,
' y ademas cada pantalla vuelve a comprobarlo en servidor (ver autorizacion.asp),
' para no repetir el fallo del legacy de "seguridad solo por ocultar el menu".
'
' "Cambiar contrasena" y "Desconectar" viven en la cabecera (header.asp), junto al
' usuario conectado, porque son acciones sobre la propia cuenta, no secciones funcionales.
Dim accionesSesion, perfilSesion, rutaActual
accionesSesion = Session("AccionesPermitidas")
perfilSesion = Session("CodPerfil")
rutaActual = Request.ServerVariables("SCRIPT_NAME")

Function EsSeccionActiva(carpeta)
    EsSeccionActiva = (InStr(1, rutaActual, "/modules/" & carpeta & "/", 1) = 1)
End Function

Function ClaseSiActiva(carpeta)
    If EsSeccionActiva(carpeta) Then
        ClaseSiActiva = " class=""activo"""
    Else
        ClaseSiActiva = ""
    End If
End Function

' Codigos leidos de Application (cargados en global.asa): 0 queries de BD por pagina.
Dim puedeSolicitudes, puedeFicheros, puedeIncidencias, puedeAdmin, puedeUsuarios, puedeBloqueos
puedeSolicitudes = TienePermiso(accionesSesion, CodigoAccionCacheado("LECTURA_SOLICITUDES"))
puedeFicheros    = TienePermiso(accionesSesion, CodigoAccionCacheado("LECTURA_FICHEROS"))
puedeIncidencias = TienePermiso(accionesSesion, CodigoAccionCacheado("LECTURA_INCIDENCIAS"))
puedeAdmin       = TienePermiso(accionesSesion, CodigoAccionCacheado("ADMINISTRADOR"))
puedeUsuarios    = (perfilSesion = CodigoPerfilCacheado("Seguridad"))
puedeBloqueos    = (perfilSesion = CodigoPerfilCacheado("Administrador") Or perfilSesion = CodigoPerfilCacheado("Tramitación"))
%>
<nav class="menu-lateral">
<ul>
    <li<%= ClaseSiActiva("home") %>>
        <a href="/modules/home/home.asp">
            <svg class="menu-icono" xmlns="http://www.w3.org/2000/svg" width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <path d="M3 9l9-7 9 7v11a2 2 0 01-2 2H5a2 2 0 01-2-2z"/><polyline points="9 22 9 12 15 12 15 22"/>
            </svg>
            Inicio
        </a>
    </li>

    <% If puedeSolicitudes Or puedeIncidencias Or puedeFicheros Or puedeBloqueos Then %>
    <li class="menu-grupo-label" aria-hidden="true">Gestión</li>
    <% End If %>

    <% If puedeSolicitudes Then %>
    <li<%= ClaseSiActiva("solicitudes") %>>
        <a href="/modules/solicitudes/consulta.asp">
            <svg class="menu-icono" xmlns="http://www.w3.org/2000/svg" width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <path d="M16 4h2a2 2 0 012 2v14a2 2 0 01-2 2H6a2 2 0 01-2-2V6a2 2 0 012-2h2"/><rect x="8" y="2" width="8" height="4" rx="1"/>
            </svg>
            Solicitudes
        </a>
    </li>
    <% End If %>

    <% If puedeIncidencias Then %>
    <li<%= ClaseSiActiva("incidencias") %>>
        <a href="/modules/incidencias/buscar.asp">
            <svg class="menu-icono" xmlns="http://www.w3.org/2000/svg" width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <path d="M10.29 3.86L1.82 18a2 2 0 001.71 3h16.94a2 2 0 001.71-3L13.71 3.86a2 2 0 00-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>
            </svg>
            Incidencias
        </a>
    </li>
    <% End If %>

    <% If puedeFicheros Then %>
    <li<%= ClaseSiActiva("ficheros") %>>
        <a href="/modules/ficheros/buscar.asp">
            <svg class="menu-icono" xmlns="http://www.w3.org/2000/svg" width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <path d="M22 19a2 2 0 01-2 2H4a2 2 0 01-2-2V5a2 2 0 012-2h5l2 3h9a2 2 0 012 2z"/>
            </svg>
            Ficheros
        </a>
    </li>
    <% End If %>

    <% If puedeBloqueos Then %>
    <li<%= ClaseSiActiva("bloqueos") %>>
        <a href="/modules/bloqueos/bloqueo.asp">
            <svg class="menu-icono" xmlns="http://www.w3.org/2000/svg" width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <rect x="3" y="11" width="18" height="11" rx="2"/><path d="M7 11V7a5 5 0 0110 0v4"/>
            </svg>
            Bloqueos
        </a>
    </li>
    <% End If %>

    <% If puedeUsuarios Or puedeAdmin Then %>
    <li class="menu-grupo-label" aria-hidden="true">Administración</li>
    <% End If %>

    <% If puedeUsuarios Then %>
    <li<%= ClaseSiActiva("usuarios") %>>
        <a href="/modules/usuarios/listar.asp">
            <svg class="menu-icono" xmlns="http://www.w3.org/2000/svg" width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <path d="M17 21v-2a4 4 0 00-4-4H5a4 4 0 00-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 00-3-3.87"/><path d="M16 3.13a4 4 0 010 7.75"/>
            </svg>
            Usuarios
        </a>
    </li>
    <% End If %>

    <% If puedeAdmin Then %>
    <li<%= ClaseSiActiva("admin") %>>
        <a href="/modules/admin/parametros.asp">
            <svg class="menu-icono" xmlns="http://www.w3.org/2000/svg" width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
                <circle cx="12" cy="12" r="3"/><path d="M19.4 15a1.65 1.65 0 00.33 1.82l.06.06a2 2 0 010 2.83 2 2 0 01-2.83 0l-.06-.06a1.65 1.65 0 00-1.82-.33 1.65 1.65 0 00-1 1.51V21a2 2 0 01-4 0v-.09A1.65 1.65 0 009 19.4a1.65 1.65 0 00-1.82.33l-.06.06a2 2 0 01-2.83-2.83l.06-.06A1.65 1.65 0 004.68 15a1.65 1.65 0 00-1.51-1H3a2 2 0 010-4h.09A1.65 1.65 0 004.6 9a1.65 1.65 0 00-.33-1.82l-.06-.06a2 2 0 012.83-2.83l.06.06A1.65 1.65 0 009 4.68a1.65 1.65 0 001-1.51V3a2 2 0 014 0v.09a1.65 1.65 0 001 1.51 1.65 1.65 0 001.82-.33l.06-.06a2 2 0 012.83 2.83l-.06.06A1.65 1.65 0 0019.4 9a1.65 1.65 0 001.51 1H21a2 2 0 010 4h-.09a1.65 1.65 0 00-1.51 1z"/>
            </svg>
            Parámetros
        </a>
    </li>
    <% End If %>
</ul>
</nav>
