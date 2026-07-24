# 5. Arquitectura ASP clásico — ToperBTR

Estado: cerrado a nivel de diseño de alto nivel y de las piezas críticas (hash, IDs, validación, catálogos, transacciones). Pendiente solo de decisiones de detalle no bloqueantes (proveedor ADO/ODBC concreto, nº de iteraciones PBKDF2).

## Capas

```
Presentación (.asp)  →  Controladores (.asp, dispatch por cmd)  →  Negocio (BLL includes)
        →  Acceso a datos (DAL includes, siempre parametrizado)
Transversal: Seguridad/Autorización · Auditoría · Utilidades comunes
```

## Estructura de carpetas

```
/app
 ├─ global.asa                    Session_OnStart/OnEnd (sin caché de aplicación en Application_OnStart)
 ├─ /include
 │   ├─ /common
 │   │   ├─ validacion.asp        Módulo centralizado: regex por tipo (IMEI, MSISDN, fecha, texto libre+maxlength)
 │   │   ├─ fechas.asp            Utilidades de fecha (sustituye el patrón día/mes/año fragmentado del legacy)
 │   │   └─ errores.asp           Página/función de error genérica, sin trazas técnicas al usuario
 │   ├─ /data
 │   │   ├─ conexion.asp          Apertura/cierre ADO/ODBC MySQL, BeginTrans/CommitTrans/RollbackTrans
 │   │   ├─ catalogos.asp         Resolución de accion_ref/perfil_ref/estado_*_ref/tipo_*_ref por NOMBRE, sin caché
 │   │   └─ secuencias.asp        Helper para leer LAST_INSERT_ID() en la misma conexión tras cada INSERT
 │   ├─ /security
 │   │   ├─ crypto_pbkdf2.asp     <script language="JScript" runat="server"> — SHA-256/HMAC/PBKDF2 nativo
 │   │   ├─ auth.asp              Login, verificación de hash (por usuarios.algoritmo_password), bloqueo por intentos fallidos, rehash progresivo
 │   │   ├─ autorizacion.asp      Comprobación unificada de permisos (sustituye los 3 mecanismos mezclados del legacy — ver 04-matriz-permisos.md)
 │   │   ├─ csrf.asp              Token de sesión por formulario (el legacy no tiene ninguno)
 │   │   └─ auditoria.asp         Registro de acciones sensibles (nunca contraseñas/hashes)
 │   └─ /layout
 │       └─ header.asp / footer.asp / menu.asp   (Tiles nunca se usó de verdad en el legacy — aquí el layout es real)
 ├─ /modules
 │   └─ /home /login /usuarios /solicitudes /incidencias /ficheros /bloqueos /admin
 ├─ /dal                          Un fichero .asp por entidad, SIEMPRE parametrizado (ADO Command+Parameters)
 ├─ /assets  (css/js/img)
 └─ /errors
```

## Decisiones ya cerradas y cómo entran en la arquitectura

- **Hash de contraseña**: `crypto_pbkdf2.asp`. Formato autodescriptivo `PBKDF2SHA256$<iteraciones>$<salt_base64>$<hash_base64>`. Columna `usuarios.algoritmo_password` discrimina legacy (`SHA1B64`) vs. nuevo. Rehash progresivo y transparente en el primer login exitoso.
- **IDs**: `AUTO_INCREMENT` en las 9 tablas afectadas (usuarios, solicitud, cliente, lista_negra, lista_blanca, historico_bloqueo_directo, ficheros, incidencia, comentario_incidencia) + `secuencias.asp` para leer `LAST_INSERT_ID()` en la misma conexión/transacción justo después del INSERT. Los códigos de negocio (`COD_SOLICITUD`, `COD_INCIDENCIA` a 6 dígitos) se calculan **después** de conocer el ID generado, no antes.
- **Catálogos**: `catalogos.asp`, sin caché de aplicación, consulta directa en cada request. Resolución siempre por nombre/descripción, nunca literal numérico hardcodeado (corrige el patrón de "magic numbers" del legacy).
- **Validación**: `validacion.asp` centralizado. Regex sobre string (nunca parseo a entero — el legacy tenía una regla de Struts Validator, nunca ejecutada en la práctica, que habría rechazado todo IMEI real por este motivo). Sin restringir el juego de caracteres de las contraseñas. `maxlength` explícito en todo texto libre, alineado a las columnas de BD.
- **Transacciones**: todo INSERT/UPDATE multi-paso envuelto en `BeginTrans`/`CommitTrans`/`RollbackTrans` real — corrige el patrón "pasos sueltos"/"rollback sobre conexión sin transacción" confirmado en varios puntos del legacy (alta de usuario, bloqueo/desbloqueo directo, cambio de estado de solicitud individual).
- **Autorización**: unificada en `autorizacion.asp`, gate real en servidor para Usuarios (perfil 5) y Bloqueos (perfiles 1 y 2) — el legacy solo los protegía ocultando el menú (ver hallazgo de seguridad en `04-matriz-permisos.md`).
- **Seguridad transversal**: sesión `HttpOnly`+`Secure`+`SameSite`; CSRF real (el legacy no tenía ninguno); auditoría de acciones sensibles sin datos secretos.

## Equivalencia Struts → ASP clásico

| Struts | ASP clásico |
|---|---|
| Action (`actions.ActionXxx`) | Controlador `.asp` con `dispatch` por `Request("cmd")` |
| ActionForm | `validacion.asp` + función de mapeo de `Request.Form`/`Request.QueryString` |
| Forward | `Response.Redirect` (patrón Post/Redirect/Get) |
| JSP + Tiles (nunca compuesto de verdad — `tiles-defs.xml` confirmado vacío) | Includes de layout reales (`header.asp`/`menu.asp`/`footer.asp`) |
| DAO/Bean Java | Ficheros `.asp` en `/dal`, un procedimiento parametrizado por operación |

## Pendiente (no bloqueante)

- Proveedor ADO/ODBC concreto para MySQL (Connector/ODBC vs. alternativa OLE DB) — se resuelve al escribir `conexion.asp`.
- Nº de iteraciones de PBKDF2 — pendiente de medir en el entorno real (JScript es interpretado).
