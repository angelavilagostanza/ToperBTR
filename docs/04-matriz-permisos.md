# 4. Matriz de permisos — ToperBTR

Estado: **cerrado** (verificado contra código real y datos reales de `accion_permitida_ref` en producción, 2026-07-23).

## Mecanismos de permisos detectados en el legacy (a unificar en ToperBTR)

El legacy no usa un único mecanismo, sino tres solapados:

1. **Por acción** (`accion_permitida_ref`): gobierna Solicitudes, Ficheros, Incidencias y el flag `ADMINISTRADOR`/`SEGURIDAD` a alto nivel.
2. **Por perfil hardcodeado en código** (literal `perfil==5`, `perfil==1`): gobierna operaciones de Usuarios y el acceso al menú de Administración. **No hay ninguna fila en `accion_permitida_ref` para Usuarios ni para Bloqueos.**
3. **Visibilidad de menú** (`ActionMenu`): decide qué enlaces se muestran, sin que necesariamente coincida con lo que la acción comprobaría si se invocara directamente por URL.

### Hallazgo de seguridad confirmado (código + datos reales)

`ActionBloqueo.getPermiso()` devuelve siempre `AccionVO.SIN_CONTROL` (0 = sin restricción), y no existe ninguna fila de `accion_permitida_ref` para Bloqueos en ningún perfil. **Conclusión: hoy, cualquier usuario autenticado con cualquier perfil puede invocar `/bloqueos.do` directamente por URL y bloquear/desbloquear un IMEI**, sin que el código lo impida — lo único que lo evita en la práctica es que el menú no muestra el enlace a los perfiles que no deberían tenerlo. Mismo riesgo, aunque no verificado a nivel de código con la misma certeza, en las operaciones de Usuarios reservadas a perfil 5 (no hay fila de permiso que las respalde tampoco).

**Decisión de diseño para ToperBTR**: eliminar la "seguridad por ocultación de menú". Cada operación sensible comprueba su permiso en el servidor en cada request, no solo en la construcción del menú:
- Bloqueos (bloqueo/desbloqueo directo) → gate explícito a perfiles 1 y 2.
- Usuarios (alta/baja/desbloqueo/cambio de contraseña de terceros) → gate explícito a perfil 5.
- Solicitudes/Ficheros/Incidencias → gate por acción, reutilizando el catálogo `accion_ref` ya confirmado.

## Catálogo `accion_ref` (confirmado por código, `AccionVO.java`)

| Código | Nombre |
|---|---|
| 0 | SIN_CONTROL |
| 10 | LECTURA_SOLICITUDES |
| 11 | ESCRITURA_SOLICITUDES |
| 20 | LECTURA_FICHEROS |
| 21 | ESCRITURA_FICHEROS |
| 30 | CONSOLIDACION |
| 40 | LECTURA_INCIDENCIAS |
| 41 | ESCRITURA_INCIDENCIAS |
| 100 | ADMINISTRADOR |
| 101 | SEGURIDAD |

## Matriz real (`SELECT * FROM accion_permitida_ref`, producción, 2026-07-23)

| Perfil | Solicitudes | Ficheros | Incidencias | Admin/Seguridad |
|---|---|---|---|---|
| **1 — Administrador** | Lectura + Escritura (10, 11) | Lectura + Escritura + Consolidación (20, 21, 30) | Lectura + Escritura/comentarios (40, 41) | ADMINISTRADOR (100) |
| **2 — Tramitación** | Lectura + Escritura (10, 11) | Lectura + Consolidación (20, 30) — **sin** Escritura (21) | Lectura + Escritura/comentarios (40, 41) | — |
| **3 — Consultas** | Solo lectura (10) | Solo lectura (20) | Solo lectura (40) — **sin** capacidad de comentar (41) | — |
| **4 — Tercero** | — | Solo lectura (20) | — | — |
| **5 — Seguridad** | — | — | — | SEGURIDAD (101) |
| Usuarios (todos los perfiles) | — sin fila en `accion_permitida_ref`; gestionado por literal `perfil==5` hardcodeado en código para las operaciones de alta/baja/desbloqueo/cambio de password de terceros | | | |
| Bloqueos (todos los perfiles) | — sin fila en `accion_permitida_ref`; `ActionBloqueo` devuelve `SIN_CONTROL` — **gate real solo por ocultación de menú, ver hallazgo de seguridad arriba** | | | |

Notas de precisión sobre esta matriz:
- Corrige una suposición de un análisis anterior: el perfil 3 (Consultas) **no** puede añadir comentarios a incidencias, solo consultarlas (no tiene la acción 41).
- El perfil 1 es el único con `ESCRITURA_FICHEROS` (21) — coherente con que es el único con acceso a la pantalla de publicación de ficheros / administración técnica.
- El código de acción `101 SEGURIDAD` está asignado a perfil 5 en los datos reales, pero no se ha confirmado en el código en qué punto exacto se comprueba `usuario.puede(101)` — el enforcement real observado en código para el perfil 5 es el literal hardcodeado `perfil==5`, no una llamada genérica al catálogo de acciones. A tener en cuenta: puede que el código 101 sea en parte vestigial/decorativo.

## Pendiente

Ninguno relevante — la matriz queda cerrada con datos reales. Solo queda como nota de diseño (no de verificación) decidir si en ToperBTR se corrige la asimetría "Consolidación sin Escritura de ficheros" para perfil 2, o se mantiene igual que hoy (recomendación: mantener igual, es una distinción de negocio deliberada, no un error).
