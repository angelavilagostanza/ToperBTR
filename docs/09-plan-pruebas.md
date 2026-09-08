# Plan de pruebas funcionales — ToperBTR

Entregable 8. Pruebas a ejecutar contra `eir_test` antes de dar el visto bueno al despliegue en producción.  
Ejecutar en orden: las pruebas de Login (E1) son prerequisito para todo lo demás.

---

## Prerequisito: ejecutar script de verificación

Antes de aplicar los scripts 001-007, ejecutar:
```
db/008_verificacion_predespliegue.sql
```
Todos los resultados `ERROR` deben estar resueltos antes de continuar.

Orden de aplicación de scripts:
```
001_entregable1_login_auth.sql
002_entregable2_usuarios.sql
003_entregable3_solicitudes.sql
004_entregable4_incidencias.sql   (si existe; verificar)
005_entregable5_ficheros.sql       (si existe; verificar)
006_bloqueos.sql
007_admin.sql                      (si existe; verificar)
```

---

## E1 — Login y autenticación

| # | Acción | Resultado esperado |
|---|--------|--------------------|
| 1.1 | Login con usuario/password correctos (hash SHA1B64 legacy) | Acceso correcto, sesión iniciada |
| 1.2 | Login con usuario/password correctos (hash PBKDF2SHA256 ToperBTR) | Acceso correcto |
| 1.3 | Login con password incorrecta | Error, sin acceso. Contador de fallos incrementado en BD |
| 1.4 | Repetir `NUM_MAX_INTENTOS_FALLIDOS` veces (por defecto 5) | En el intento N, cuenta queda bloqueada (ESTADO=2 en `usuarios`) |
| 1.5 | Login con cuenta bloqueada | Error "cuenta bloqueada", sin acceso |
| 1.6 | Login con usuario inexistente | Error genérico (igual que password incorrecta, no revelar si el usuario existe) |
| 1.7 | Token CSRF manipulado en POST de login | Redirect a login con error de sesión |
| 1.8 | Acceder a `/modules/home/home.asp` sin sesión | Redirect a login |
| 1.9 | Cerrar sesión | `Session.Abandon`, redirect a login |

---

## E2 — Usuarios

Prerequisito: sesión activa con perfil **Seguridad**.

| # | Acción | Resultado esperado |
|---|--------|--------------------|
| 2.1 | Listar usuarios | Tabla con todos los usuarios activos/bloqueados del operador |
| 2.2 | Dar de alta un usuario nuevo (login único, perfil Administrador) | Usuario creado, flash de éxito, aparece en el listado |
| 2.3 | Intentar crear usuario con login ya existente | Error "login ya existe" |
| 2.4 | Intentar crear usuario superando el límite del operador | Error "número máximo alcanzado" |
| 2.5 | Cambiar contraseña (autoservicio): contraseña correcta, nueva >= 10 chars | Contraseña cambiada, exige la nueva en el próximo login |
| 2.6 | Cambiar contraseña: contraseña actual incorrecta | Error, sin cambio |
| 2.7 | Cambiar contraseña: reutilizar una de las `NUM_OLD_PASSWORD` anteriores | Error "contraseña ya usada anteriormente" |
| 2.8 | Contraseña expirada (FECHA_CADUCIDAD < NOW()) | Al login, redirect a cambio de contraseña obligatorio |
| 2.9 | Reset de contraseña por Seguridad (otro usuario) | Contraseña cambiada, audit registrado con CAMBIO_PASSWORD_ADMIN |
| 2.10 | Dar de baja un usuario | Estado → Borrado, no aparece en el listado activo |
| 2.11 | Desbloquear una cuenta bloqueada por intentos fallidos | Estado → Activo, puede volver a hacer login |
| 2.12 | Acceder a `/modules/usuarios/listar.asp` con perfil no-Seguridad | Redirect a error "sin permiso" |

---

## E3 — Solicitudes

Prerequisito: sesión activa con perfil que tenga `ESCRITURA_SOLICITUDES` (p.ej. Tramitación o Administrador) o solo `LECTURA_SOLICITUDES` para las pruebas de solo lectura.

| # | Acción | Resultado esperado |
|---|--------|--------------------|
| 3.1 | Crear solicitud de **Inclusión** con IMEI de LISTA_BLANCA | Solicitud creada, estado "En tramitación", flash de éxito |
| 3.2 | Crear solicitud de **Exclusión** para un IMEI con Inclusión existente | Solicitud de Exclusión creada, vinculada a la Inclusión |
| 3.3 | Intentar Exclusión sin Inclusión previa para ese IMEI | Error "no existe solicitud de Inclusión previa" |
| 3.4 | Intentar crear solicitud con IMEI ya en curso | Error "ya existe solicitud en curso" |
| 3.5 | Consultar solicitudes con filtros (IMEI, tipo, estado, fechas) | Resultados filtrados correctamente |
| 3.6 | Ver detalle de solicitud | Datos completos: IMEI, cliente, tipo, estado, historial |
| 3.7 | Cancelar solicitud "En tramitación" | Estado → Cancelada |
| 3.8 | Intentar cancelar solicitud ya cerrada/cancelada | Checkbox no aparece en UI; POST con id manipulado → sin efecto |
| 3.9 | Acceder sin `LECTURA_SOLICITUDES` | Redirect a "sin permiso" |

---

## E4 — Incidencias

Prerequisito: sesión activa con `LECTURA_INCIDENCIAS`.

| # | Acción | Resultado esperado |
|---|--------|--------------------|
| 4.1 | Buscar incidencias con filtros (objeto, tipo, estado, fechas) | Lista de incidencias correctamente filtrada |
| 4.2 | Ver detalle de incidencia asociada a un fichero | Detalle correcto (JOIN corregido: `TIPO_OBJETO='F'`) |
| 4.3 | Ver detalle de incidencia asociada a una solicitud | Detalle correcto (`TIPO_OBJETO='S'`) |
| 4.4 | Añadir comentario a una incidencia (con `ESCRITURA_INCIDENCIAS`) | Comentario visible en el detalle, audit registrado |
| 4.5 | Intentar añadir comentario vacío | Error de validación |
| 4.6 | Acceder sin `LECTURA_INCIDENCIAS` | Redirect a "sin permiso" |

---

## E5 — Ficheros

Prerequisito: sesión activa con `LECTURA_FICHEROS`.

| # | Acción | Resultado esperado |
|---|--------|--------------------|
| 5.1 | Buscar ficheros (sin filtros) | Lista de todos los ficheros salvo el genérico |
| 5.2 | Buscar por nombre exacto | Solo el fichero con ese nombre |
| 5.3 | Filtrar por operador (seleccionar Grupo MASMOVIL) | Ficheros con la clave del operador home en el nombre + ficheros EIR (si perfil Admin) |
| 5.4 | Descargar un fichero existente | Descarga el contenido del fichero desde la ruta en BD |
| 5.5 | Intentar descargar un fichero manipulando el nombre en URL | La ruta se lee siempre de BD (sin concatenar path del usuario), sin traversal |
| 5.6 | Descargar fichero EIR con perfil no-Administrador (URL directa) | Redirect a "sin permiso" |
| 5.7 | Acceder sin `LECTURA_FICHEROS` | Redirect a "sin permiso" |

---

## E6 — Bloqueos

Prerequisito: sesión activa con perfil **Administrador** o **Tramitación**.

| # | Acción | Resultado esperado |
|---|--------|--------------------|
| 6.1 | Bloquear IMEI de LISTA_BLANCA (cliente activo GMM) | Bloqueo insertado, flash de éxito, aparece en el listado |
| 6.2 | Intentar bloquear IMEI que no está en LISTA_BLANCA | Error "no pertenece a cliente activo de Grupo MASMOVIL" |
| 6.3 | Intentar bloquear un IMEI ya bloqueado | Error "ya se encuentra bloqueado" |
| 6.4 | Bloquear IMEI que ya tiene un bloqueo GLOBAL → tipo cambia a COMBINADO | Fila existente actualizada a BLOQUEO DIRECTO Y GLOBAL |
| 6.5 | Buscar bloqueos (filtros: IMEI, tipo, razon, fechas) | Resultados correctamente filtrados |
| 6.6 | Seleccionar IMEIs y desbloquear en lote | IMEIs eliminados de LISTA_NEGRA, audit registrado |
| 6.7 | Desbloquear IMEI sin bloqueo directo | Flash "no se pudo desbloquear ninguno" (sin bloqueo DIRECTO activo) |
| 6.8 | Acceder con perfil no-Administrador y no-Tramitación | Redirect a "sin permiso" |
| 6.9 | CSRF: enviar POST de bloqueo con token inválido | Redirect a error de sesión |

---

## E7 — Administración

Prerequisito: sesión activa con acción `ADMINISTRADOR`.

| # | Acción | Resultado esperado |
|---|--------|--------------------|
| 7.1 | Abrir `/modules/admin/parametros.asp` | Lista todos los parámetros de `PARAMETROS_REF` con valores editables |
| 7.2 | Cambiar el valor de `NUM_MAX_INTENTOS_FALLIDOS` | Guardado, flash de éxito, nuevo valor visible en recarga |
| 7.3 | Verificar que el cambio de `NUM_MAX_INTENTOS_FALLIDOS` surte efecto | Repetir test 1.4 con el nuevo valor |
| 7.4 | Intentar guardar con CLAVE manipulada (clave inexistente) | Error "el parámetro no existe" |
| 7.5 | CSRF: POST con token inválido | Redirect a error de sesión |
| 7.6 | Acceder con perfil no-Administrador | Redirect a "sin permiso" |

---

## Pruebas transversales de seguridad

| # | Acción | Resultado esperado |
|---|--------|--------------------|
| X.1 | Acceder a `/include/data/conexion.asp` directamente vía HTTP | 403 (bloqueado por `include/web.config`) |
| X.2 | Acceder a `/dal/bloqueos_dal.asp` directamente vía HTTP | 403 (bloqueado por `dal/web.config`) |
| X.3 | Acceder a `/modules/admin/parametros.asp` con perfil que no es Administrador | Redirect a "sin permiso" (gate en servidor) |
| X.4 | Salida de `Server.HTMLEncode` en campos de texto libres | No se ejecuta JS en comentarios/nombres/valores de parámetros |
| X.5 | Inyección SQL: introducir `' OR '1'='1` en cualquier campo de búsqueda | Tratar como texto literal (queries parametrizadas) |
| X.6 | Pedir un módulo al que se tiene permiso de menú pero el token de sesión expiró | Redirect a login |
| X.7 | Petición GET directa a `*_do.asp` | CSRF fallará (formulario no genera token en GET), redirect a error de sesión |
| X.8 | Verificar que `web.config` deshabilita `errorMode=Detailed` | En producción, los errores 500 muestran la página amigable, no stack traces |

---

## Auditoría — verificación post-pruebas

Tras ejecutar los tests anteriores, comprobar en `auditoria_accion`:

```sql
SELECT accion, entidad, resultado, COUNT(*) AS total
FROM auditoria_accion
GROUP BY accion, entidad, resultado
ORDER BY accion;
```

Deben aparecer registros para:
- `LOGIN_OK`, `LOGIN_FAIL`, `LOGIN_BLOQUEADO` (E1)
- `USUARIO_ALTA`, `USUARIO_BAJA`, `USUARIO_DESBLOQUEO`, `CAMBIO_PASSWORD` (E2)
- `SOLICITUD_ALTA`, `SOLICITUD_CANCELAR` (E3)
- `INCIDENCIA_COMENTARIO` (E4)
- `BLOQUEO_INSERTAR`, `BLOQUEO_QUITAR` (E6)
- `MODIFICAR` en `PARAMETROS_REF` (E7)

Ningún registro debe tener `detalle` con contraseñas ni hashes.
