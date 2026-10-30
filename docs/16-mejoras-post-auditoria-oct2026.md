# 16. Mejoras implementadas tras la auditoría de seguridad — octubre 2026

Resumen de los cambios realizados **después** de cerrar la auditoría de seguridad (ver [`docs/15-auditoria-seguridad-oct2026.md`](15-auditoria-seguridad-oct2026.md)). A diferencia de esa auditoría, lo recogido aquí no corrige vulnerabilidades: son peticiones de negocio, ajustes de usabilidad y una corrección ortográfica completa del texto de la aplicación. Periodo: 2026-10-29 a 2026-10-30. Todo lo descrito está validado en pre y **desplegado en producción**.

## Resumen

| Cambio | Alcance | Estado |
|---|---|---|
| Renombrado "Inclusión"/"Exclusión" → "Bloqueo Terminal"/"Desbloqueo Terminal" | BD + código | Desplegado |
| Campo "Enlace externo" en Solicitudes | BD + código | Desplegado |
| Timeout de sesión 20 → 60 minutos | Configuración | Desplegado |
| Texto de bienvenida sobre el EIR | Código | Desplegado |
| Enlace bidireccional Solicitudes ↔ Incidencias | Código | Desplegado |
| Corrección ortográfica completa (tildes) | Código | Desplegado |
| Consolidado de cambios de BD para la transición | Documentación/BD | Pendiente de revisión por Infraestructura |

## Renombrado de terminología: Inclusión/Exclusión → Bloqueo Terminal/Desbloqueo Terminal

Petición de negocio: sustituir en toda la aplicación el texto heredado de la nomenclatura del sistema legacy Java ("Inclusión" / "Exclusión") por un lenguaje más claro para el usuario final: **"Bloqueo Terminal"** y **"Desbloqueo Terminal"**.

El cambio tiene dos partes:
- **Base de datos**: un único `UPDATE` sobre la tabla de catálogo de tipos de solicitud, identificando cada fila por su clave interna (no por el texto que se está reemplazando). Esa misma tabla alimenta directamente el desplegable de alta de una solicitud, el listado de solicitudes y el detalle de cada una — así que esta única actualización propaga el nuevo texto a los tres sitios sin necesidad de tocar nada más.
- **Código**: el título de la página de alta, los mensajes de validación de ese formulario, y el mensaje de error que aparece al intentar crear un desbloqueo sin una solicitud de bloqueo previa sobre el mismo IMEI.

No se ha tocado, porque es un concepto distinto, la "fecha de inclusión"/"fecha de exclusión" que aparece en el módulo de Bloqueos: ahí se refiere a cuándo un IMEI entra o sale de la lista negra de terminales bloqueados, no al tipo de solicitud que se está renombrando aquí.

## Nuevo campo "Enlace externo" en Solicitudes

Petición de negocio: poder guardar, junto a cada solicitud, un enlace a un sistema externo (Jira, un CRM, etc.) relacionado con esa gestión.

- Campo de texto libre (hasta 500 caracteres) en el formulario de alta de una solicitud, con el texto de ayuda "Enlace a Jira, ID en CRM, etc..".
- Se muestra en el detalle de la solicitud: si el valor introducido es una dirección web completa, aparece como un enlace en el que se puede hacer clic; si es un identificador corto (por ejemplo, el código de un ticket), se muestra como texto simple.
- El campo es opcional y **no se puede modificar una vez creada la solicitud** — no existe hoy ninguna pantalla para editar una solicitud ya existente en la aplicación, y se decidió no construir una solo para este caso.

## Ampliación del tiempo de sesión: de 20 a 60 minutos

Varios usuarios reportaron que la sesión caducaba con demasiada frecuencia durante el uso normal de la aplicación, obligándoles a volver a iniciar sesión. Se ha ampliado el tiempo de inactividad permitido antes de que la sesión expire de 20 a 60 minutos.

## Ajuste del texto de bienvenida sobre el sistema EIR

A petición de negocio, se ha simplificado la frase introductoria de la pantalla de inicio que explica qué es el EIR (Equipment Identity Register), el registro donde se incluye el IMEI de un terminal para bloquearlo.

## Enlace entre Solicitudes e Incidencias

Antes de este cambio, al consultar una incidencia vinculada a una solicitud ya se podía navegar directamente a esa solicitud — pero el camino contrario no existía: al ver el detalle de una solicitud con incidencias asociadas, solo se mostraba un contador ("Esta solicitud tiene 2 incidencia(s) asociada(s)"), sin ninguna forma de ver cuáles eran.

Se ha añadido ese enlace que faltaba: desde el detalle de una solicitud se puede acceder directamente al listado de las incidencias relacionadas con ella. De paso, también se ha añadido un enlace directo a la solicitud correspondiente en el propio listado de incidencias, donde antes solo se mostraba el código como texto.

## Corrección ortográfica completa de la aplicación

Se realizó una revisión sistemática de todo el texto que ve el usuario en la aplicación (títulos, etiquetas de formulario, mensajes de error y de confirmación, cabeceras de tabla) en busca de palabras mal escritas, principalmente tildes que faltaban — un patrón que se arrastraba desde el desarrollo inicial del proyecto. Se encontraron y corrigieron 121 casos repartidos en 28 páginas distintas de todos los módulos de la aplicación (Usuarios, Solicitudes, Incidencias, Bloqueos, Ficheros, Administración, Auditoría, inicio de sesión y página de inicio).

Durante esta revisión se detectó y resolvió un problema técnico de base: ninguna de las páginas de la aplicación tenía declarada su codificación de caracteres de forma explícita, lo que en ciertas condiciones podía hacer que los acentos y la letra "ñ" no se mostraran correctamente en el navegador. Se ha corregido añadiendo esa declaración en todas las páginas afectadas, sin que esto cambie nada más allá de garantizar que el texto se vea siempre tal como está escrito.

## Consolidado de cambios de base de datos para la transición

Se ha preparado un fichero único (`db/013_Modificacion_BD_Transicion.sql`) que reúne en un solo lugar todos los cambios de base de datos que ToperBTR necesita aplicar sobre el entorno de producción definitivo, pensado para que el equipo de Infraestructura pueda revisar y valorar el impacto de cada cambio antes de ejecutarlo. Incluye tanto los cambios históricos de los distintos entregables como los descritos en este mismo documento (el renombrado de tipos de solicitud y el nuevo campo de enlace externo). Está pendiente de esa revisión; no se ha aplicado todavía contra el entorno de producción definitivo.
