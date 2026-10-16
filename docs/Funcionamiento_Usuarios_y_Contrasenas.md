# Funcionamiento de Usuarios y Contraseñas — ToperBTR

Documento de referencia sobre cómo funciona el ciclo de vida de cuentas de usuario y contraseñas en ToperBTR, extraído directamente del código (`app/`), no de los documentos de diseño originales — estos pueden haber quedado desactualizados respecto a la implementación final. Todas las afirmaciones citan el fichero y la línea de origen. Creado 2026-10-06; actualizado 2026-10-17 (reescritura de la criptografía a VBScript puro, y corrección de los hallazgos de seguridad A1/A2 sobre el bloqueo por intentos fallidos).

## 1. Hash de contraseñas

- **Algoritmo**: PBKDF2-HMAC-SHA256, implementación 100% nativa en VBScript puro (sin DLL/COM externos). Ver `app/include/security/crypto_pbkdf2.asp`.
  - Hasta 2026-10-14 estaba en JScript; se reescribió por completo en VBScript tras descubrir que el servidor de producción (Windows Server 2025) crashea de forma intermitente al mezclar motores VBScript+JScript en la misma página ASP — no es un problema de este código en particular, sino de todo el proyecto, resuelto eliminando JScript de raíz.
- **Iteraciones**: 200, fijadas en `Application("PBKDF2Iteraciones")` (`app/global.asa:7`) — no es un parámetro de `parametros_ref`, es de configuración de aplicación.
  - Algunos hashes antiguos se crearon con 50.000 iteraciones (valor previo, antes de bajarlo a 200); `login_do.asp` amplía `Server.ScriptTimeout` a 360s precisamente para poder verificarlos y migrarlos.
  - **Nota de seguridad aceptada (2026-10-17)**: 200 iteraciones es un valor bajo frente a lo que recomiendan hoy OWASP/NIST para PBKDF2-HMAC-SHA256 (cientos de miles). Subirlo de forma realista es inviable con esta implementación en VBScript puro (≈6 ms por iteración; 10.000 iteraciones tardarían ≈1 minuto por login). Decisión consciente: se asume el riesgo residual, compensado por los controles de BD y acceso que gestiona el equipo de Infraestructura, en vez de complicar el código o degradar la experiencia de login.
- **Sal**: 16 bytes aleatorios por contraseña (`bytesAleatorios(16)`). No es criptográficamente fuerte (`Rnd()` de VBScript), pero es aceptable porque la sal solo necesita ser única, no secreta.
- **Formato de almacenamiento**: `PBKDF2SHA256$<iteraciones>$<sal_base64>$<hash_base64>`, clave derivada de 32 bytes. Comparación en tiempo constante (`ComparacionSegura`).
- **Convivencia con el legacy**: la columna `usuarios.algoritmo_password` (`VARCHAR(20) DEFAULT 'SHA1B64'`) discrimina el formato por usuario. `VerificarPassword()` usa PBKDF2 si el algoritmo es `PBKDF2SHA256`; si no, cae al legacy **SHA-1 + Base64 sin sal**, que se mantiene solo para poder verificar y migrar cuentas antiguas, nunca para crear hashes nuevos.
- **Migración automática (re-hash on login)**: en cada login correcto, si el algoritmo almacenado no es `PBKDF2SHA256`, o lo es pero con un número de iteraciones distinto al actual, se recalcula el hash y se persiste ya en PBKDF2/200 iteraciones (`app/include/security/auth.asp`). También se migra en cualquier cambio de contraseña (propio o por admin), ya que `ActualizarPassword` siempre escribe `ALGORITMO_PASSWORD = 'PBKDF2SHA256'`.
- Las cuentas creadas desde ToperBTR (alta de usuario) nacen siempre en PBKDF2SHA256, nunca en el formato legacy.

## 2. Requisitos de la contraseña

**Única regla: longitud mínima de 10 caracteres.** No hay ninguna exigencia de mayúsculas, minúsculas, dígitos ni símbolos, en ningún punto del código.

Se valida de forma idéntica (longitud ≥10 + confirmación coincidente) en los tres flujos que crean o cambian una contraseña:
- Alta de usuario (`app/modules/usuarios/alta_do.asp`)
- Cambio propio (`app/modules/usuarios/cambio_password_do.asp`)
- Reseteo por administrador (`app/modules/usuarios/reset_password_admin_do.asp`)

El `minlength="10"` del formulario de reseteo por admin es solo una ayuda de cliente; el servidor siempre re-valida.

## 3. Histórico de contraseñas / no reutilización

Punto de entrada único para ambos flujos de cambio: `CambiarPasswordConHistorico` (`app/dal/usuarios_dal.asp`).

- **No se puede reutilizar** la contraseña actual ni ninguna de las guardadas en el histórico (`PasswordReutilizada`, compara contra hashes tanto legacy como PBKDF2 según corresponda por fila).
- **Profundidad del histórico**: parámetro `NUM_OLD_PASSWORD` (por defecto 5 si no está definido en `parametros_ref`). Mientras haya menos de ese número de filas se inserta una nueva; al llegar al máximo, se recicla la fila con `FECHA_CADUCO` más antigua en vez de seguir creciendo.
- **Aplica igual** a cambio propio y a reseteo por administrador — decisión explícita de ToperBTR para corregir la asimetría del legacy, donde el reseteo por admin no comprobaba el histórico.
- Rotación de histórico y actualización de la contraseña van en la misma transacción (`BeginTrans`/`CommitTrans`/`RollbackTrans`), así que un fallo revierte ambas operaciones.

## 4. Caducidad de la contraseña

- **Duración**: parámetro `DURACION_PASSWORD`, por defecto 60 días si no está definido. Se recalcula en cada alta y en cada cambio de contraseña (propio o por admin): `FECHA_CADUCIDAD = Now() + DURACION_PASSWORD días`.
- **Al caducar**: el login se rechaza directamente con *"Su contraseña ha caducado. Contacte con el administrador para renovarla."* — **no existe ningún flujo de autoservicio** para cambiar una contraseña caducada desde la pantalla de login. La única salida es que un administrador de perfil Seguridad se la resetee.
- ⚠️ **Inconsistencia detectada**: la descripción del parámetro en Administración → Parámetros dice *"0 = sin caducidad"*, pero el código no trata ese valor como caso especial — con `DURACION_PASSWORD=0` la cuenta quedaría con `FECHA_CADUCIDAD` prácticamente en el instante de creación, es decir, caducada casi de inmediato, no "sin caducidad". Pendiente de decidir: corregir el código para que `0` deshabilite la comprobación, o corregir el texto de la UI.

## 5. Bloqueo automático por intentos fallidos

Toda la lógica vive en `app/include/security/auth.asp`.

- **Umbral**: parámetro `NUM_MAX_INTENTOS_FALLIDOS`, configurado a **5** (sembrado explícitamente en `db/001_entregable1_login_auth.sql` y `db/009_reaplicacion_idempotente.sql`; el mismo valor es el *fallback* del código si el parámetro no existiera).
- **Contador**: `usuarios.num_fallos`, se incrementa en cada login fallido por contraseña incorrecta. Al alcanzar el umbral, la cuenta pasa a `ESTADO = Bloqueado` (excepto si ya está `Borrada`) y se audita como `LOGIN_BLOQUEO_AUTOMATICO`.
  - **Incremento atómico (corregido 2026-10-17)**: `RegistrarFalloLogin` incrementa con `UPDATE ... SET NUM_FALLOS = COALESCE(NUM_FALLOS, 0) + 1` directamente en SQL (con el bloqueo de fila que aplica MySQL durante el `UPDATE`), y relee el valor ya actualizado antes de decidir el bloqueo. Antes se calculaba `numFallosActual + 1` en VBScript a partir de un valor leído al principio de la petición — bajo peticiones en paralelo (p. ej. un ataque con varios hilos simultáneos), el contador real nunca avanzaba al ritmo de los intentos y el bloqueo se podía evitar por completo. El comportamiento que ve el usuario normal (mensajes, umbral, bloqueo) no cambia; lo que cambia es que ahora es imposible burlarlo con peticiones simultáneas.
- **Mensajes progresivos** durante los intentos:
  - Quedan 3 o más intentos → mensaje genérico "Usuario o contraseña incorrectos."
  - Quedan 2 intentos → avisa explícitamente.
  - Queda 1 intento → avisa explícitamente.
  - Se agota el último intento → mensaje de cuenta bloqueada.
- **Cuenta ya bloqueada**: el login se rechaza *antes* de comprobar la contraseña, con el mismo mensaje de bloqueo.
- El contador `num_fallos` se resetea a 0 en: login correcto, cualquier cambio de contraseña (propio o por admin) y desbloqueo manual.

## 5bis. Límite de intentos por IP (nuevo, 2026-10-17)

El bloqueo de la sección anterior es **por cuenta**: no impedía que alguien probase muchas cuentas distintas (o la misma, con contraseñas distintas) desde la misma conexión sin que ninguna cuenta en concreto llegase nunca a su propio umbral. Para cubrir ese hueco, `login_do.asp` comprueba un segundo límite, independiente del anterior, **antes que cualquier otra cosa** (incluso antes de validar el token del formulario):

- **Límite**: 20 intentos de login fallidos (por cualquier motivo: usuario inexistente, contraseña incorrecta, cuenta bloqueada o caducada) desde la misma dirección IP, en una ventana de 5 minutos. Valores fijos en código (`RATE_LIMIT_IP_MAX_INTENTOS`/`RATE_LIMIT_IP_VENTANA_MINUTOS`, `app/include/security/auth.asp`), no son parámetros de `parametros_ref`.
- **Al superarlo**: el login se rechaza con *"Demasiados intentos desde esta conexión. Espere unos minutos e inténtelo de nuevo."*, sin ni siquiera llegar a comprobar usuario/contraseña.
- **Se reinicia solo**, pasados los 5 minutos desde el primer fallo contado en la ventana actual — no hay desbloqueo manual para esto, ni falta hace (no es un estado de cuenta, es un límite de tráfico).
- **Almacenamiento**: en memoria del propio proceso de IIS (`Application`), no en base de datos — se pierde si el Application Pool recicla (aceptable para este propósito: en el peor caso, un atacante gana una ventana de 5 minutos extra tras un reciclado, que no es un evento frecuente ni predecible por fuera).
- Solo cuenta intentos **fallidos**, nunca los que terminan en login correcto, para no penalizar el uso normal (p. ej. varias personas detrás de la misma IP de oficina).

## 6. Desbloqueo de cuentas

- **Único camino**: acción manual "Desbloquear seleccionados" en Usuarios (perfil Seguridad), que llama a `DesbloquearUsuariosLote` — pone `ESTADO = Activo` y `NUM_FALLOS = 0`.
- **No existe** desbloqueo automático por tiempo ni autoservicio de ningún tipo.
- Resetear la contraseña de una cuenta bloqueada **no la desbloquea** — son dos acciones independientes; hay que aplicar las dos si se quiere que el usuario vuelva a entrar.
- `DesbloquearUsuariosLote` excluye explícitamente a las cuentas en estado `Borrado`: una baja lógica nunca se puede "revivir" por esta vía.

## 7. Estados de cuenta

En uso real en el código actual, exactamente tres (resueltos siempre por nombre vía `CodigoEstadoUsuario`, nunca por código numérico fijo en el código de ToperBTR):

| Estado | Cómo se llega | Efecto en login | ¿Reversible? |
|---|---|---|---|
| **Activo** | Alta de usuario, o desbloqueo manual | Permite login (sujeto a contraseña y caducidad) | — |
| **Bloqueado** | Automático tras 5 fallos, o bloqueo manual por admin | Rechazado siempre, sin comprobar contraseña | Sí, solo por desbloqueo manual |
| **Borrado** | Acción manual "Dar de baja" | Rechazado (mensaje genérico, no revela el motivo) | No, es permanente. Tampoco se puede reutilizar ese login para una cuenta nueva |

El estado legacy `LOGEADO` (valor 3 en el Java original) **no se usa en ToperBTR**: en el código legacy era solo una marca de sesión activa en memoria, no un estado de cuenta administrable, y no aparece en ningún punto del código actual de ToperBTR.

## 8. Qué comprueba el login, en orden

0. **(2026-10-17) Límite de intentos por IP** (sección 5bis) — se comprueba antes que nada, incluso antes del token CSRF del formulario.
1. El usuario existe.
2. No está en estado `Borrado`.
3. No está en estado `Bloqueado`.
4. La contraseña es correcta (si falla, incrementa el contador de fallos —y el de la IP— y puede bloquear la cuenta en este mismo paso).
5. La contraseña no ha caducado.
6. Éxito: resetea el contador de fallos, actualiza fecha de última conexión, re-hashea la contraseña si el algoritmo/iteraciones no son los actuales, y construye la sesión (perfil, operador, acciones permitidas).

⚠️ **Importante**: una vez iniciada la sesión, las páginas posteriores solo comprueban `Session("Autenticado") = True` — **no se vuelve a comprobar el estado de la cuenta ni la caducidad de la contraseña durante la sesión activa**. Si un administrador bloquea o da de baja a un usuario mientras este tiene sesión abierta, puede seguir operando hasta que la sesión caduque por inactividad (20 minutos, `Application("SessionTimeoutMinutos")` en `global.asa`) o cierre sesión voluntariamente.

## 9. Reseteo de contraseña por el administrador

- El administrador **escribe directamente** la nueva contraseña en el formulario — no se genera automáticamente ni se envía por correo.
- **No necesita conocer la contraseña actual** del usuario (a diferencia del cambio propio, que sí la exige).
- Se aplican las mismas reglas de longitud (≥10) e histórico/no-reutilización que en el cambio propio.
- **No obliga a cambiarla en el siguiente login** — la contraseña puesta por el admin queda activa de inmediato, con su nueva fecha de caducidad (no existe ningún flag de "debe cambiar la contraseña" en el esquema).
- Se audita como `CAMBIO_PASSWORD_ADMIN`, distinto de `CAMBIO_PASSWORD` (cambio propio).

## 10. Valores por defecto al crear un usuario

- `ESTADO`: Activo.
- `NUM_FALLOS`: 0.
- `FECHA_ULTIMA_CONEXION`: nula.
- `FECHA_ALTA`: fecha/hora de creación.
- `FECHA_CADUCIDAD`: fecha de creación + `DURACION_PASSWORD` días.
- Contraseña siempre creada en PBKDF2SHA256.
- Validaciones previas: login ≤15 caracteres y no vacío, nombre/apellidos obligatorios, perfil obligatorio, contraseña válida y confirmada, login no usado antes (ni siquiera por una cuenta ya dada de baja), y que el operador no haya superado su cupo máximo de usuarios.

## Resumen: parametrizado vs. fijo en código

| Valor | Origen | Valor actual |
|---|---|---|
| Iteraciones PBKDF2 | `Application("PBKDF2Iteraciones")`, fijo en `global.asa` | 200 |
| Longitud mínima de contraseña | Literal `10` repetido en 3 ficheros `.asp` | 10 caracteres, sin reglas de complejidad |
| Profundidad del histórico | `parametros_ref.NUM_OLD_PASSWORD` | 5 (valor por defecto del código si no está en BD) |
| Duración de la contraseña | `parametros_ref.DURACION_PASSWORD` | 60 días (valor por defecto del código si no está en BD) |
| Umbral de bloqueo por fallos (por cuenta) | `parametros_ref.NUM_MAX_INTENTOS_FALLIDOS` | 5 (sembrado explícitamente en BD) |
| Límite de intentos por IP | Literal `RATE_LIMIT_IP_MAX_INTENTOS` en `auth.asp` | 20 intentos fallidos cada 5 minutos |
| Timeout de sesión | `Application("SessionTimeoutMinutos")`, fijo en `global.asa` | 20 minutos |

## Nota de seguridad ya resuelta

Existía un fichero `app/modules/admin/reset_hash_test.asp`, marcado en su propio código como *"UTILIDAD TEMPORAL - ELIMINAR DESPUÉS DE USAR"*, que permitía resetear la contraseña de cualquier usuario a un valor conocido sin CSRF, sin sesión, sin comprobación de perfil, sin reglas de longitud/histórico y sin auditoría — la única protección era una lista fija de IPs permitidas. Se saltaba todas las reglas descritas en este documento. **Eliminado del repositorio el 2026-10-06.**
