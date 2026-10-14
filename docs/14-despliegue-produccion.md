# 14. Procedimiento de despliegue a Producción — ToperBTR

Documento de referencia para subir ToperBTR al entorno de producción real (no `eir_test`/`toperbtr-pre`). Construido a partir de lo ya decidido en `docs/07-plan-fases.md`/`docs/08-cambios-bbdd.md`, de lo confirmado en esta conversación, y de una revisión completa del estado actual del código y los scripts de BD.

## Decisiones ya tomadas (esta conversación, 2026-10-13)

1. **El sitio IIS de PRO ya existe** — este despliegue es de código, no de infraestructura nueva.
2. **La migración de BD real se coordina con DBA/infra, con backup previo** — no se aplica nada contra `eir` de producción sin ventana acordada ni copia de seguridad previa.
3. **El endpoint de intercambio con operadores (`app/externos/intercambio/`) queda fuera de esta primera ola** — depende de coordinación externa (URL con los operadores) y de infraestructura (certificado de cliente en IIS) que no está lista; va como workstream aparte con su propia fecha.

## La "mejor opción": una sola entrega de código, puesta en marcha progresiva

`docs/07-plan-fases.md` ya decidió el orden correcto (cambios de BD aditivos primero, verificar que el legacy sigue intacto, luego ToperBTR en convivencia) — no hay que inventar nada ahí. Dado que los 8 entregables ya están completos, no tiene sentido un despliegue de código "módulo a módulo" (ToperBTR es una aplicación aparte del legacy, no se puede activar media app). La recomendación concreta:

1. **BD primero, sola.** Aplicar los cambios aditivos a `eir` real, verificar que el legacy sigue funcionando exactamente igual, **antes** de tocar IIS de PRO.
2. **Código de ToperBTR completo, de una vez**, pero con **arranque progresivo de usuarios**: en vez de anunciar la URL de PRO a todo el mundo el primer día, dar acceso primero a un grupo pequeño (p. ej. el propio equipo de Seguridad/Administrador) durante unos días, con el legacy como respaldo para el resto, antes de generalizar. Así el radio de impacto de cualquier sorpresa es pequeño, sin necesitar ningún mecanismo de "feature flag" dentro del código (que no existe ni hace falta construir).
3. **Intercambio con operadores, aparte**, como ya se ha decidido.

El resto de este documento desarrolla esos tres pasos.

## Bloqueantes a cerrar antes de empezar

| # | Bloqueante | Estado | Acción |
|---|---|---|---|
| 1 | Duración del periodo de convivencia y criterio de "cuándo apagar" el legacy | **Sin definir con negocio** (`docs/07-plan-fases.md`) | No bloquea el primer despliegue, pero hay que fijarlo antes de prometer una fecha de apagado del legacy |
| 2 | Ningún script de `db/` se ha ejecutado nunca contra `eir` real — todos están escritos y probados solo contra `eir_test`, varios con `USE eir_test;` literal | Confirmado por inspección directa de `db/*.sql` | Ver Fase 1 — hay que preparar una versión apta para producción |
| 3 | Estado real de `eir_test` (si los scripts 001-010 están aplicados) sin confirmación reciente | Sin confirmar desde 2026-09-14 | Re-verificar con `db/008_verificacion_predespliegue.sql` antes de dar nada por bueno, aunque sea solo para no arrastrar una base de referencia incorrecta a producción |
| 4 | `app/conn_test.asp` sigue trackeado en git y desplegable — página de diagnóstico de esquema de BD **sin ningún tipo de autenticación** | Confirmado, sigue en el repo hoy | **Eliminarlo antes de desplegar a PRO** (mismo criterio que ya se aplicó a `reset_hash_test.asp`) |
| 5 | Bug real en `ContarIncidencias` (join sin discriminador `TIPO_OBJETO`, inconsistente con `BuscarIncidencias`/`ObtenerIncidenciaPorIndice`) | **Corregido en esta sesión** (2026-10-13) | Ninguna — ya resuelto |
| 6 | IMEIs reales subidos por error al repo de git (`FicherosBTR/intercambio/*`) | **Resuelto** (2026-10-01): `.gitignore`, sacados del tracking, historial purgado con `filter-branch` + `push --force-with-lease`, repo confirmado privado | Ninguna — ya resuelto, pese a que la documentación histórica (`docs/11`) todavía lo liste como pendiente |
| 7 | Formato nuevo de `GenerarCodigoSolicitud` sin confirmar con negocio/BA (compatibilidad con sistemas externos que consumen el código de solicitud) | Sin resolver | Confirmar antes de ir a producción si algún sistema externo consume ese código |
| 8 | `app/global.asa` de producción: hay que construirlo **a mano a partir del `global.asa` actual**, nunca copiando `global.asa.example` | El `.example` no tiene `Sub CargarCatalogosCacheados` ni la llamada en `Application_OnStart` — usarlo tal cual dejaría `CodigoAccionCacheado`/`CodigoPerfilCacheado` siempre en `-1`, que **deniega todo acceso** en el arranque | Ver Fase 3 |
| 9 | Modo de pipeline de IIS (Clásico vs. Integrado) en el sitio de PRO | Sin confirmar | `app/include/web.config` solo protege de verdad en modo Integrado; confirmar con quien administre el sitio, o añadir autorización manual por IIS Manager si es Clásico |
| 10 | Cabecera `Strict-Transport-Security` deshabilitada a propósito en `app/web.config` (comentario: "activar solo cuando SSL esté instalado en producción") | Pendiente de SSL | Si el sitio de PRO ya tiene HTTPS, descomentar la cabecera como parte del despliegue (ver Fase 3) |

## Fase 1 — Preparar los scripts de BD para producción real

Ningún script existente apunta a `eir` real; hay que montar uno nuevo, consolidado, siguiendo el mismo patrón idempotente que ya probó su valía en `db/009_reaplicacion_idempotente.sql`:

1. **No asumir nada sobre el estado de `eir` real** a partir de lo que se sepa de `eir_test` — son bases de datos distintas, y `eir` lleva corriendo con el legacy en producción sin que se le haya tocado nada de esto. Primer paso real: ejecutar contra `eir` un diagnóstico de solo lectura equivalente al de `db/008_verificacion_predespliegue.sql` para saber exactamente qué falta.
2. Construir `db/012_despliegue_produccion.sql` (nombre propuesto) que:
   - Consolide todo lo de `db/001` a `db/010` en un único script idempotente (mismo patrón de `db/009`: `ALTER...MODIFY COLUMN` es idempotente de por sí en MySQL; `ADD COLUMN`/`INSERT` protegidos con comprobación previa de `information_schema`/`WHERE NOT EXISTS`).
   - **Sin ningún `USE eir_test;` hardcodeado** — que lo seleccione quien lo ejecute, contra la base de datos correcta.
   - Incluya los mismos puntos de verificación previa que ya señala `docs/08-cambios-bbdd.md` (valores reales de `parametros_ref`, si ya existe `estado_usuario_ref` fila Bloqueado, colisiones de ID antes de fijar `AUTO_INCREMENT`).
3. Revisar si `005`/`007` (referenciados como "si existe" en `docs/09-plan-pruebas.md`) hacen falta de verdad: Ficheros (Entregable 5) es de solo lectura sin cambios de esquema, y Admin (Entregable 7) solo edita filas ya existentes de `parametros_ref` — con toda probabilidad nunca hizo falta un script para ninguno de los dos. Confirmarlo y, si es así, quitar esa referencia confusa del plan de pruebas.

## Fase 2 — Ventana de migración coordinada con DBA/infra

Según lo acordado:

1. **Backup completo de `eir`** antes de tocar nada.
2. Ventana coordinada con quien administre la BD de producción, avisando también a quien mantenga la app legacy (sigue funcionando contra la misma base durante y después del cambio).
3. Ejecutar `db/012_despliegue_produccion.sql` (Fase 1).
4. Verificar inmediatamente después que **el legacy sigue funcionando exactamente igual** (login, bloqueo, solicitudes) — es el propio criterio ya fijado en `docs/07-plan-fases.md`, antes de seguir con el código de ToperBTR.
5. Tener claro el plan de reversión: todos los cambios son aditivos (columnas más anchas, columnas/filas nuevas, `AUTO_INCREMENT`) — revertir un `ALTER...MODIFY COLUMN` a un ancho mayor no debería hacer falta nunca (no rompe nada existente), pero documentar explícitamente con el DBA qué se haría si algo fallara a mitad de la ventana.

## Fase 3 — Desplegar el código en el sitio IIS de PRO

1. **Quitar `app/conn_test.asp`** antes de copiar nada (bloqueante #4).
2. Copiar `app/`, `db/` (de referencia, no se ejecuta como parte del deploy de IIS) a la ruta física del sitio de PRO.
3. **Construir el `global.asa` de PRO a mano**, partiendo del `global.asa` actual (no del `.example`):
   - `DBConnectionString`: host/credenciales reales de `eir` de producción.
   - Mantener `PBKDF2Iteraciones = 200` (no el `50000` que todavía aparece en `.example`, que es un valor viejo/obsoleto).
   - Confirmar que `Sub CargarCatalogosCacheados` y su llamada en `Application_OnStart` están presentes (bloqueante #8).
4. Confirmar modo de pipeline de IIS (bloqueante #9) en el sitio de PRO.
5. Si el sitio de PRO ya tiene SSL/HTTPS instalado, descomentar la cabecera `Strict-Transport-Security` en `app/web.config` (bloqueante #10) como parte de este mismo despliegue.
6. Confirmar que el directorio virtual IIS `ficherosBTR` (usado por `ficheros/descargar.asp`) apunta a la ruta física correcta en el servidor de PRO.
7. **No desplegar `app/externos/`** en esta ola (decisión ya tomada) — o, si se copia por comodidad junto al resto, dejar su ruta sin publicar/sin DNS hasta que el workstream de operadores esté listo.

## Fase 4 — Validación post-despliegue

Reutilizar `docs/09-plan-pruebas.md` como base, pero ejecutándolo contra el entorno de PRO real:

1. Login + bloqueo por intentos fallidos (con un usuario de prueba, nunca con una cuenta real).
2. Flujo completo de Usuarios, Solicitudes, Incidencias, Ficheros, Bloqueos, Admin.
3. Caso concreto a probar explícitamente: una incidencia vinculada a un **fichero** (no solo a una solicitud) — para confirmar en real el fix de `TIPO_OBJETO` (bloqueante #5, ya corregido en código, pero nunca verificado con datos reales).
4. Prueba X.8 de `docs/09`: confirmar que un error 500 real muestra la página amigable, no un stack trace, en el entorno de PRO.
5. Verificar en `AUDITORIA_ACCION` que las acciones de esta misma validación quedan registradas (ver `docs/13-auditoria-acciones.md`).

## Fase 5 — Arranque progresivo y convivencia

1. Dar acceso primero a un grupo reducido (equipo de Seguridad/Administrador) antes de anunciar la URL de PRO a todos los usuarios.
2. Durante la convivencia, el legacy sigue disponible y apuntando a la misma BD — cualquier incidencia con ToperBTR tiene vuelta atrás inmediata sin tocar datos (el usuario simplemente vuelve a usar el legacy).
3. Pendiente de fijar con negocio (bloqueante #1): duración de la convivencia y criterio de apagado definitivo del legacy. No bloquea el despliegue, pero conviene cerrarlo antes de comunicar fechas.

## Explícitamente fuera de esta ola

- **`app/externos/intercambio/`** (descarga de ficheros de intercambio EIR para otros operadores): requiere coordinar con los operadores la URL definitiva y con infra el certificado de cliente en IIS — ver `docs/12-endpoint-intercambio-operadores.md`. Va como workstream propio, con su propia validación antes de publicarse.

## Checklist resumen (orden de ejecución)

- [ ] Confirmar con negocio el formato de `GenerarCodigoSolicitud` (bloqueante #7), si aplica.
- [ ] Eliminar `app/conn_test.asp` del repo.
- [ ] Re-verificar estado de `eir_test` con `db/008` (por higiene, no bloquea PRO pero evita arrastrar una referencia incorrecta).
- [ ] Construir `db/012_despliegue_produccion.sql` (consolidado, idempotente, sin `USE eir_test`).
- [ ] Diagnóstico de solo lectura contra `eir` real, antes de asumir nada.
- [ ] Backup de `eir` + ventana coordinada con DBA/infra.
- [ ] Ejecutar `db/012` contra `eir` real.
- [ ] Verificar que el legacy sigue funcionando igual.
- [ ] Construir `global.asa` de PRO a mano (a partir del actual, no del `.example`).
- [ ] Confirmar modo de pipeline IIS y, si hay SSL, activar HSTS.
- [ ] Desplegar el código (sin `app/externos/`) al sitio de PRO.
- [ ] Ejecutar la validación de la Fase 4, incluido el caso de incidencia vinculada a fichero.
- [ ] Arranque progresivo: grupo reducido primero, luego generalizar.
- [ ] Fijar con negocio duración de convivencia y criterio de apagado del legacy.
