# 13. Auditoría de acciones — ToperBTR

Documento de referencia: qué se registra en la tabla de auditoría, cómo se escribe y cómo consultarlo. Extraído directamente del código actual (`app/`), citando fichero y línea para cada punto.

## La tabla: `AUDITORIA_ACCION`

Tabla nueva de ToperBTR (aditiva, sin relación con el esquema legacy), creada en `db/001_entregable1_login_auth.sql`:

```sql
CREATE TABLE auditoria_accion (
  id INT NOT NULL AUTO_INCREMENT,
  cod_indice_usuario INT NULL,
  accion VARCHAR(50) NOT NULL,
  entidad VARCHAR(50) NOT NULL,
  entidad_id VARCHAR(50) NULL,
  fecha DATETIME NOT NULL,
  ip VARCHAR(45) NULL,
  resultado VARCHAR(20) NOT NULL,
  detalle VARCHAR(500) NULL,
  PRIMARY KEY (id),
  KEY idx_auditoria_usuario (cod_indice_usuario),
  KEY idx_auditoria_fecha (fecha)
)
```

| Columna | Contenido |
|---|---|
| `cod_indice_usuario` | Usuario que realizó la acción. **Nulo** cuando no hay usuario identificado (login fallido con usuario inexistente, o acciones del endpoint de intercambio con operadores externos, que no tiene sesión). |
| `accion` | Literal de texto que identifica el hecho (`LOGIN`, `USUARIO_ALTA`, `BLOQUEO_INSERTAR`, etc. — ver tabla completa más abajo). No hay catálogo/tabla de referencia para estos valores, son literales definidos en cada punto del código. |
| `entidad` | Tabla/dominio de negocio afectado (`USUARIOS`, `SOLICITUD`, `LISTA_NEGRA`, `FICHEROS`, `INCIDENCIA`, `PARAMETROS_REF`). |
| `entidad_id` | Identificador del registro afectado (puede ser numérico o texto según el caso — p. ej. el IMEI en Bloqueos, la clave del parámetro en Administración). Nulo cuando no aplica. |
| `fecha` | Marca de tiempo del servidor (`NOW()`). |
| `ip` | IP de origen de la petición, capturada automáticamente. |
| `resultado` | Código corto del resultado (`OK`, `FALLO`, `BLOQUEADO`, `NO_ENCONTRADO`, etc. — varía según la acción). |
| `detalle` | Texto libre con contexto adicional (máx. 500 caracteres). **Nunca contiene contraseñas ni hashes** — es una regla explícita del diseño. |

## Cómo se escribe

Punto único de escritura: `Sub RegistrarAuditoria(conn, codIndiceUsuario, accion, entidad, entidadId, resultadoTexto, detalle)`, en `app/include/security/auditoria.asp`.

- La **IP se captura sola** (`Request.ServerVariables("REMOTE_ADDR")`) — no hay que pasarla.
- Es **"best effort"**: va envuelta en `On Error Resume Next` y usa la variante de base de datos que no redirige ante un fallo (`EjecutarNonQueryTx`). Si el propio `INSERT` de auditoría fallara, **nunca** bloquea ni interrumpe la operación de negocio real que se está auditando.
- `codIndiceUsuario` acepta `Null` explícitamente para acciones sin usuario identificado.

## Acciones registradas hoy

| `ACCION` | `ENTIDAD` | Cuándo se genera | `RESULTADO` posibles | Origen |
|---|---|---|---|---|
| `LOGIN` | USUARIOS | Cada intento de inicio de sesión, correcto o fallido (con el motivo del fallo) | `OK`, `FALLO` (detalle: "Usuario no encontrado", "Usuario dado de baja", "Cuenta bloqueada", "Contrasena incorrecta", "Password caducada") | `app/include/security/auth.asp` |
| `LOGIN_BLOQUEO_AUTOMATICO` | USUARIOS | El propio intento de login que hace que la cuenta supere el umbral de intentos fallidos y pase a Bloqueada | `BLOQUEADO` | `app/include/security/auth.asp` |
| `LOGOUT` | USUARIOS | Cierre de sesión explícito | `OK` | `app/modules/login/logout.asp` |
| `USUARIO_ALTA` | USUARIOS | Alta de un usuario nuevo | `OK` (detalle: login creado) | `app/modules/usuarios/alta_do.asp` |
| `USUARIO_MODIFICACION` | USUARIOS | Edición de nombre/apellidos/perfil de un usuario existente | `OK` | `app/modules/usuarios/editar_do.asp` |
| `USUARIO_BAJA` | USUARIOS | Acción masiva "Dar de baja" sobre usuarios seleccionados (una fila por usuario) | `OK` | `app/modules/usuarios/accion_ejecutar.asp` |
| `USUARIO_BLOQUEO` | USUARIOS | Acción masiva "Bloquear" (bloqueo manual por admin, distinto del automático por intentos fallidos) | `OK` | `app/modules/usuarios/accion_ejecutar.asp` |
| `USUARIO_DESBLOQUEO` | USUARIOS | Acción masiva "Desbloquear" | `OK` | `app/modules/usuarios/accion_ejecutar.asp` |
| `CAMBIO_PASSWORD` | USUARIOS | Cambio de contraseña por el propio usuario (autoservicio) | `OK` | `app/modules/usuarios/cambio_password_do.asp` |
| `CAMBIO_PASSWORD_ADMIN` | USUARIOS | Reseteo de contraseña de un usuario por un administrador (perfil Seguridad) | `OK` | `app/modules/usuarios/reset_password_admin_do.asp` |
| `SOLICITUD_ALTA` | SOLICITUD | Alta de una solicitud de Inclusión o Exclusión | `OK` (detalle: código de solicitud) | `app/modules/solicitudes/alta_do.asp` |
| `SOLICITUD_CANCELAR` | SOLICITUD | Cancelación de una solicitud (una fila por solicitud cancelada en el lote) | `OK` | `app/modules/solicitudes/cancelar_ejecutar.asp` |
| `INCIDENCIA_COMENTARIO` | INCIDENCIA | Alta de un comentario sobre una incidencia | `OK` | `app/modules/incidencias/comentario_do.asp` |
| `BLOQUEO_INSERTAR` | LISTA_NEGRA | Bloqueo directo de un IMEI (`entidad_id` = el IMEI) | `OK`, `YA_BLOQUEADO`, `NO_YOIGO` | `app/modules/bloqueos/bloquear_do.asp` |
| `BLOQUEO_QUITAR` | LISTA_NEGRA | Desbloqueo de IMEIs en lote (una sola fila por lote completo, con el recuento en `detalle`) | `OK` | `app/modules/bloqueos/desbloquear_do.asp` |
| `FICHERO_DESCARGA` | FICHEROS | Descarga de un fichero desde la pantalla de administración (usuario autenticado) | `OK`, `NO_ENCONTRADO`, `FICHERO_FISICO_NO_ENCONTRADO` | `app/modules/ficheros/descargar.asp` |
| `FICHERO_DESCARGA_INTERCAMBIO` | FICHEROS | Descarga desde el endpoint externo para operadores (`app/externos/intercambio/descargar.asp`), sin sesión — se audita **también lo rechazado**, es la única traza de quién usa ese endpoint | `OK`, `NO_ENCONTRADO`, `RECHAZADO_FORMATO` | `app/externos/intercambio/descargar.asp` |
| `MODIFICAR` | PARAMETROS_REF | Cambio de un parámetro del sistema (`entidad_id` = la clave del parámetro) | `OK` (detalle: valor anterior y nuevo) | `app/modules/admin/parametro_do.asp` |

## Qué NO se audita (conocido, no es un olvido)

- **Nunca** se guarda contraseña ni hash, en ningún caso — es una regla de diseño explícita, no una omisión.
- No hay alta manual de incidencias en la aplicación (las crea el sistema/EIRControl), solo comentarios — por eso no existe un `INCIDENCIA_ALTA`.
- La sesión no se vuelve a comprobar en cada página tras el login — si un admin bloquea o da de baja a un usuario con sesión abierta, esa persona sigue operando hasta que la sesión caduque; esas operaciones posteriores **sí** quedan auditadas bajo su acción propia, pero no hay un evento específico de "sesión invalidada a mitad de uso".

## Cómo consultarlo

Módulo **Auditoría** (`app/modules/auditoria/buscar.asp`), gateado al perfil **Seguridad** (mismo perfil que Usuarios). Permite buscar por Usuario, Acción (desplegable con los valores de la tabla anterior en uso real) y rango de fechas; el campo de texto libre sobre **Detalle** solo se aplica si ya hay otro criterio informado (evita un `LIKE` sin acotar sobre una columna sin índice).
