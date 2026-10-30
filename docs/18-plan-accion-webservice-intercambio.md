# 18. Plan de acción — BTR-WS (webservice de intercambio para operadores)

Plan de arranque para la próxima sesión de trabajo. Ver diseño completo en [`docs/17-arquitectura-webservice-intercambio.md`](17-arquitectura-webservice-intercambio.md).

## Día 1 — Preguntas que hay que cerrar antes de escribir código

No bloquean crear el esqueleto del proyecto, pero sí bloquean el diseño fino de la validación de IP y del *binding* de IIS:

1. ¿El middleware que valida el certificado de cliente reenvía la IP real del operador en alguna cabecera (`X-Forwarded-For` u otra), o IIS solo verá la IP del propio middleware?
2. ¿El tramo middleware → IIS va por HTTPS interno o es red interna sin TLS?
3. Confirmar con Infraestructura: ¿sitio IIS nuevo con *binding* propio para `btrws.yoigo.com`, o aplicación nueva dentro de un sitio existente pero en pool distinto?
4. ¿Está instalado ya en el servidor el **ASP.NET Core Hosting Bundle (.NET 8)**? (`dotnet --info` o comprobar en "Programas y características" → "Microsoft ASP.NET Core 8.0 Hosting Bundle").
5. Nombre definitivo de repositorio/carpeta de despliegue (propuesto: `BTRWS`).

## Paso 1 — Esqueleto del proyecto (sin esperar a las respuestas de arriba)

- [ ] Crear repositorio nuevo `BTRWS` (separado de `ToperBTR`).
- [ ] `dotnet new webapi -minimal` sobre .NET 8 LTS, estructura según `docs/17-...md`.
- [ ] `Program.cs`: pipeline con orden explícito de middleware:
  1. Rate limiting (`AddRateLimiter`, built-in de ASP.NET Core).
  2. Lista blanca de IP (middleware propio, lee de BD con caché corta en memoria — ver Paso 3).
  3. Endpoint `GET /intercambio/descargar`.
- [ ] `appsettings.json` sin secretos: la cadena de conexión ODBC va solo en `appsettings.Production.json` (no versionado) o en variable de entorno del pool de IIS.
- [ ] Reutilizar la misma validación de nombre de fichero que ya existe en el `.asp` actual (1 letra de operador + 8 dígitos) — mismo contrato de negocio, nueva implementación en C#.

## Paso 2 — Migración de base de datos

- [ ] Crear `db/014_whitelist_ip_webservice.sql` en el repo de **ToperBTR** (la tabla vive en la BD compartida, no en el repo nuevo) con el `CREATE TABLE ws_ip_permitida` de `docs/17-...md`.
- [ ] Aplicarlo primero en pre (recordar: pre y PRO comparten la misma BD física `eir_test` en `172.30.170.27` — un único `CREATE TABLE` ya vale para ambos).
- [ ] Insertar manualmente, vía SQL directo, las primeras IPs conocidas de operadores para poder probar en Paso 4 (hasta que exista la pantalla de administración).

## Paso 3 — Lógica de la aplicación

- [ ] `IpAllowListService`: consulta `ws_ip_permitida` (solo filas `activo=1`), con una caché en memoria de corta duración (p.ej. 60s) para no golpear la BD en cada petición — el riesgo de servir con una lista ligeramente desfasada 60s es aceptable frente al de sobrecargar la BD compartida con ToperBTR.
- [ ] Middleware de IP: compara `HttpContext.Connection.RemoteIpAddress` (o la cabecera que confirme la pregunta 1 del Día 1) contra la lista. Si no coincide → `403` + auditoría, sin llegar al endpoint.
- [ ] `FicheroIntercambioService`: misma validación de formato + comprobación de operador existente que el `.asp` actual, more el acceso físico al fichero (mismo *share*/ruta de ficheros de intercambio que usa ToperBTR hoy — confirmar ruta exacta con el equipo antes de codificar).
- [ ] `AuditoriaService`: escribe en la **misma tabla `auditoria_accion`** que ya usa ToperBTR (coherencia: toda la auditoría del ecosistema EIR en un solo sitio), con una `ACCION` nueva y distinta, p.ej. `FICHERO_DESCARGA_WS` — sin FK a `usuarios` (la petición no tiene usuario autenticado humano).
- [ ] Logging estructurado (Event Log o fichero) de cualquier rechazo (IP no permitida, formato inválido, fichero no encontrado) — igual de exhaustivo que la auditoría en BD, como segunda traza independiente.

## Paso 4 — Pruebas locales (antes de tocar IIS)

- [ ] `dotnet run` en local contra la BD de pre: probar IP no incluida (403), fichero con formato inválido (400), operador inexistente (400), fichero válido pero no existente en BD (404), fichero real existente (200 + binario correcto).
- [ ] Confirmar que **ninguna** respuesta de error devuelve HTML ni traza de stack — todo texto plano / JSON de error mínimo.

## Paso 5 — Despliegue en IIS (depende de respuestas del Día 1)

- [ ] Instalar el Hosting Bundle si falta (pregunta 4) + `iisreset`.
- [ ] Crear pool `BTRWS_AppPool` ("No Managed Code", identidad dedicada, `AlwaysRunning`).
- [ ] Crear sitio/aplicación nueva con el *binding* acordado (pregunta 3).
- [ ] `dotnet publish` + copiar al servidor, en carpeta física separada de `V:\ToperBTR`.
- [ ] Verificar `web.config` generado (`hostingModel="InProcess"`) y que no hereda ningún `httpErrors` de otro sitio.
- [ ] Coordinar con el equipo de red el *routing* desde `btrws.yoigo.com` → middleware → este nuevo *binding*.
- [ ] Repetir las pruebas del Paso 4 ya en pre, pero a través de la URL real (de extremo a extremo, incluyendo el middleware de certificado si ya está disponible en ese entorno).

## Paso 6 — Coexistencia y retirada del endpoint `.asp` actual

- [ ] Documentar fecha de arranque de BTR-WS.
- [ ] Coordinar con operadores la migración de su integración a la nueva URL.
- [ ] Periodo de coexistencia (ambos endpoints activos, monitorizando auditoría para ver cuándo deja de usarse el antiguo).
- [ ] Retirar `app/externos/intercambio/descargar.asp` + su `web.config` + el *binding* IIS correspondiente, en un commit dedicado, solo cuando se confirme tráfico cero en el antiguo durante un periodo razonable (a definir, p.ej. 30 días sin peticiones `FICHERO_DESCARGA_INTERCAMBIO`).

## Paso 7 (futuro, fuera de este plan) — Pantalla de administración de IPs en ToperBTR

Pendiente para una iteración posterior, una vez BTR-WS esté en producción: pantalla en `toperbtr.masmovil.com` (módulo Admin existente, mismo patrón que `admin/parametros.asp`) para gestionar `ws_ip_permitida` (alta/baja/activar-desactivar) sin tocar SQL a mano. No se aborda ahora para no mezclar el arranque del webservice nuevo con una funcionalidad de UI adicional.
