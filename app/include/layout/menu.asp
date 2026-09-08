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
    <li<%= ClaseSiActiva("home") %>><a href="/modules/home/home.asp">Inicio</a></li>
    <% If puedeUsuarios Then %><li<%= ClaseSiActiva("usuarios") %>><a href="/modules/usuarios/listar.asp">Usuarios</a></li><% End If %>
    <% If puedeSolicitudes Then %><li<%= ClaseSiActiva("solicitudes") %>><a href="/modules/solicitudes/consulta.asp">Solicitudes</a></li><% End If %>
    <% If puedeIncidencias Then %><li<%= ClaseSiActiva("incidencias") %>><a href="/modules/incidencias/buscar.asp">Incidencias</a></li><% End If %>
    <% If puedeFicheros Then %><li<%= ClaseSiActiva("ficheros") %>><a href="/modules/ficheros/buscar.asp">Ficheros</a></li><% End If %>
    <% If puedeBloqueos Then %><li<%= ClaseSiActiva("bloqueos") %>><a href="/modules/bloqueos/bloqueo.asp">Bloqueos</a></li><% End If %>
    <% If puedeAdmin Then %><li<%= ClaseSiActiva("admin") %>><a href="/modules/admin/parametros.asp">Administracion</a></li><% End If %>
</ul>
</nav>
