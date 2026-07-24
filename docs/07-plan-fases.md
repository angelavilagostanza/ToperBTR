# 7. Plan por fases — ToperBTR

Estado: cerrado a nivel de alcance por entregable. Incorpora la convivencia confirmada entre la app legacy y ToperBTR sobre la misma BD de producción (sin migración de datos).

| Entregable | Objetivo | Incluye (actualizado con hallazgos del código real) | Riesgos conocidos a cubrir |
|---|---|---|---|
| **1. Base + Login + Autenticación + Autorización** | Layout, sesión, login, menú dinámico | `crypto_pbkdf2.asp`, `auth.asp` con bloqueo por intentos fallidos (nuevo — en el legacy estaba roto/desactivado desde 2011), rehash progresivo SHA1→PBKDF2, `autorizacion.asp` unificado (ver `04-matriz-permisos.md`), `csrf.asp`, `catalogos.asp` sin caché | Pieza más nueva respecto al legacy — UAT más cuidadoso. Umbral de intentos fallidos = 5, con avisos progresivos desde el 3er fallo |
| **2. Usuarios** | Consulta/alta/baja/desbloqueo/cambio password | `validacion.asp` en formularios; no reutilización de `cod_usuario` tras baja (desviación intencionada del legacy); histórico de 5 contraseñas simétrico en autoservicio y admin; gate explícito a perfil 5 | Reproducir siempre el patrón "por lote/índice" (el único que funciona en el legacy), nunca el "por cod_usuario individual" (roto) |
| **3. Solicitudes** | Alta, consulta, detalle, cancelación | AUTO_INCREMENT + `LAST_INSERT_ID()`; consultas 100% parametrizadas (corrige la inyección SQL confirmada en `listarSolicitudes` y otras); estado vigente por `FECHA_FIN IS NULL` (no `MAX(cod_estado)`); resolución de tipo/estado siempre por catálogo, nunca literal | Módulo con más inyección SQL confirmada del legacy — máxima prioridad en revisión de queries |
| **4. Incidencias** | Consulta, detalle, comentarios | Join corregido y discriminado por `TIPO_OBJETO` (bug real confirmado en el legacy); código de negocio `COD_INCIDENCIA` a 6 dígitos; corrección opcional del bug de "todas las incidencias" (punto y coma colgante) si se decide | Validar con negocio el comportamiento nuevo del join antes de cerrar el módulo |
| **5. Ficheros** | Búsqueda, listado, descarga | Control de sesión real en la descarga (el legacy tiene un bypass de autenticación confirmado y explotable aquí); exclusión de `cod_fichero=0` | Hallazgo de seguridad más grave del legacy — máxima prioridad |
| **6. Bloqueos** | Bloqueo/desbloqueo directo, consulta | Transacciones reales (el legacy no envuelve estas operaciones en ninguna); historización simétrica en exclusión combinada (ya decidido); gate explícito a perfiles 1 y 2 (el legacy no lo tenía, solo ocultaba el menú) | Condición de carrera TOCTOU confirmada en el legacy — cubierta con transacción + AUTO_INCREMENT |
| **7. Administración técnica** | Refresco de catálogos, generación de fichero | Se simplifica respecto al legacy: no hay "refrescar caché" porque no hay caché de aplicación. Sigue marcado como opcional por el propio negocio | Prioridad baja |
| **8. Hardening + pruebas + despliegue** | Checklist IIS, pruebas de regresión, corte | Plan de despliegue en dos tiempos dado que hay convivencia confirmada: (a) aplicar los cambios de esquema aditivos sobre la BD de producción mientras el legacy sigue vivo, verificando que no le afecta; (b) periodo de convivencia con ambas apps activas sobre la misma BD; (c) apagado de la app legacy pasado ese periodo | Falta fijar con negocio la duración del periodo de convivencia y el criterio de "cuándo apagar" definitivamente el legacy |

## Orden de despliegue de BBDD (por la convivencia)

1. Aplicar todos los cambios "imprescindibles" de `08-cambios-bbdd.md` (aditivos, no rompen el legacy).
2. Verificar que la app legacy sigue funcionando igual tras aplicarlos.
3. Desplegar los módulos de ToperBTR uno a uno según el orden de entregables de arriba, con la app legacy siempre disponible como referencia/fallback durante la convivencia.
4. Apagar la app legacy cuando negocio confirme el fin del periodo de convivencia.
