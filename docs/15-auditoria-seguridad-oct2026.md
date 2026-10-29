# 15. Auditoría de seguridad — octubre 2026

Informe resumido de la auditoría de seguridad realizada sobre el código de `app/` (lo que corre hoy en producción), y de las acciones de mitigación aplicadas. Periodo: 2026-10-15 a 2026-10-29. Todos los hallazgos se corrigieron en la rama `main`, se validaron primero en el entorno de pre y después se desplegaron a producción (`https://toperbtr.masmovil.com`), salvo los que se documentan explícitamente como "riesgo aceptado".

## Resumen

| Severidad | Hallazgos | Estado |
|---|---|---|
| Crítico | 1 (C1) | Riesgo aceptado (decisión de negocio) |
| Alto | 5 (A1-A5) | Todos corregidos y desplegados |
| Medio | 7 (M1-M7) | 4 corregidos y desplegados; 3 riesgo aceptado (dependencias externas / decisión de negocio) |
| Bajo | 5 (B1-B5) | Todos corregidos y desplegados |

## Crítico

### C1 — PBKDF2 con 200 iteraciones

**Hallazgo:** el hash de contraseñas (PBKDF2-HMAC-SHA256) usa 200 iteraciones, muy por debajo de las recomendaciones actuales de OWASP/NIST (decenas o cientos de miles). Si la base de datos se filtrara, los hashes serían crackeables mucho más rápido de lo deseable.

**Por qué no se sube el número:** la implementación de la criptografía está en VBScript puro (ver más abajo, motivo técnico no relacionado con este hallazgo) y es interpretada, no compilada — cada iteración cuesta ~6 ms. Subir a un número de iteraciones realmente seguro (decenas de miles) haría que cada login tardase decenas de segundos o minutos, inviable en la práctica.

**Decisión:** riesgo aceptado. La base de datos tiene sus propias medidas de seguridad, gestionadas por el equipo de Infraestructura, que son las que de verdad mitigan este escenario (acceso restringido, cifrado en reposo, etc.).

## Alto

### A1 — Condición de carrera en el bloqueo de cuentas por intentos fallidos

**Hallazgo:** el contador de intentos fallidos de login se leía al principio de la petición y se escribía como "valor leído + 1" al final. Bajo peticiones paralelas (un ataque de fuerza bruta con varios hilos simultáneos), varias peticiones podían leer el mismo valor inicial y el contador real nunca avanzaba al ritmo de los intentos — el bloqueo automático de la cuenta se podía evitar por completo.

**Mitigación:** el incremento del contador se hace ahora de forma atómica en SQL (`NUM_FALLOS = COALESCE(NUM_FALLOS, 0) + 1`), aprovechando el bloqueo de fila que aplica MySQL durante el `UPDATE`, y se relee el valor ya actualizado antes de decidir si se bloquea la cuenta.

### A2 — Sin límite de intentos de login por IP

**Hallazgo:** el único control de fuerza bruta era por cuenta (A1). Nada impedía probar muchas cuentas distintas desde la misma IP sin que ninguna en concreto llegara a su propio umbral de bloqueo.

**Mitigación:** nuevo límite de 20 intentos fallidos por IP cada 5 minutos, independiente del bloqueo por cuenta, implementado en memoria de aplicación de IIS (sin tocar base de datos).

### A3 — El perfil "Tercero" podía ver/descargar ficheros de otros operadores

**Hallazgo:** el perfil "Tercero" (pensado para usuarios de otros operadores de telefonía, no de MASMOVIL/Yoigo) tiene permiso de lectura de Ficheros, el mismo que varios perfiles internos a los que sí les corresponde ver los ficheros de todos los operadores. El filtro de operador en el buscador era de libre elección incluso para este perfil, y la descarga no comprobaba a qué operador pertenecía el fichero solicitado — un usuario Tercero legítimo podía ver/descargar ficheros de intercambio EIR de otros operadores.

**Nota:** en el momento de la corrección no había ningún usuario con perfil Tercero dado de alta en el sistema; se corrigió como prevención.

**Mitigación:** el filtro de operador se fuerza ahora en servidor al propio operador del usuario cuando el perfil es Tercero (ignorando cualquier valor recibido del navegador), y la descarga comprueba explícitamente que el fichero solicitado pertenece a ese operador antes de servirlo.

### A4 — XSS reflejado en el listado de usuarios

**Hallazgo:** los filtros de perfil y estado en la paginación del listado de usuarios se insertaban en los enlaces de la página sin codificar, a diferencia del resto de filtros de la misma pantalla. Permitía construir una URL que, al ser abierta por un administrador de Seguridad, ejecutaba JavaScript arbitrario en su navegador.

**Mitigación:** codificación de URL añadida a esos dos filtros, igual que ya tenía el resto.

### A5 — Cookie de sesión sin protección HttpOnly/Secure

**Hallazgo:** ASP clásico no permite marcar como `HttpOnly` la cookie de sesión automática de IIS, por lo que seguía siendo legible por JavaScript — un XSS (como A4) podría robarla y suplantar la sesión de otro usuario.

**Mitigación:** se añadió una segunda cookie de "huella de sesión", esta sí con el atributo `HttpOnly` (y `Secure` cuando la conexión es HTTPS), generada en el login y verificada en cada petición protegida. Un atacante que robe la cookie de sesión original no puede leer ni reenviar esta segunda cookie, así que no puede suplantar la sesión con ella sola.

## Medio

### M1 — XSS almacenado vía el login de un usuario

**Hallazgo:** el listado de usuarios escribía el login de cada usuario dentro de un comentario HTML de depuración sin codificar. Combinado con que el alta de usuarios no restringía qué caracteres podía contener el login, un administrador de Seguridad podía crear una cuenta con un login diseñado para ejecutar JavaScript en el navegador de cualquier otro administrador que abriera ese listado.

**Mitigación:** eliminado el comentario de depuración; el login de un usuario nuevo debe ajustarse ahora a un formato que excluye los caracteres que permitirían este ataque (solo letras, números y los símbolos habituales de un email/login).

### M2 — Fuga de detalles técnicos de la base de datos al crear una solicitud

**Hallazgo:** si fallaba el alta de una solicitud, el mensaje de error mostrado al usuario incluía el detalle técnico crudo devuelto por el motor de base de datos (tipo de error, nombres reales de tablas y restricciones internas) — información que facilita a un atacante entender la estructura interna del sistema.

**Mitigación:** ese detalle técnico se registra ahora en el log interno del servidor; al usuario se le muestra siempre el mismo mensaje genérico que ya usa el resto de la aplicación en estos casos.

### M3 — Generador de aleatoriedad no criptográfico

**Hallazgo:** los tokens de seguridad (protección contra CSRF, huella de sesión de A5, sal de las contraseñas) se generaban con el generador de números aleatorios estándar de VBScript, que no está diseñado para fines de seguridad y es, en teoría, predecible.

**Mitigación:** se sustituyó por un generador que usa el proveedor criptográfico del propio sistema operativo Windows, tras comprobar en el servidor real qué mecanismos estaban disponibles y verificar cuál de ellos funcionaba correctamente desde este entorno.

### M4 — Faltaban cabeceras de seguridad HTTP

**Hallazgo:** no había ninguna política de seguridad de contenido (CSP) configurada, y la cabecera que fuerza el uso de HTTPS (HSTS) estaba preparada pero desactivada a la espera de que el sitio tuviera certificado SSL.

**Mitigación:** activada la cabecera HSTS (el sitio ya tiene HTTPS) y añadida una política de seguridad de contenido que bloquea la carga de scripts, estilos y otros recursos desde dominios ajenos a la propia aplicación, reduciendo el impacto de cualquier intento de inyección de código que pudiera aparecer en el futuro.

### M5 — Mensajes de login distintos según el motivo del fallo

**Hallazgo:** el sistema de login devuelve un mensaje distinto cuando una cuenta está bloqueada o cuando la contraseña ha caducado, frente al mensaje genérico para credenciales incorrectas — lo que en teoría permite a alguien averiguar qué cuentas existen y en qué estado están, sin necesidad de acertar ninguna contraseña.

**Decisión:** riesgo aceptado. ToperBTR es una aplicación de uso interno, no expuesta a Internet; en ese contexto, quién tiene cuenta no es información sensible (ya es visible por otros medios internos), por lo que unificar los mensajes solo perjudicaría a los usuarios legítimos sin reducir ningún riesgo real.

### M6 — Conexión a la base de datos sin cifrado TLS explícito

**Hallazgo:** la cadena de conexión a MySQL no fuerza explícitamente el uso de TLS.

**Decisión:** riesgo aceptado como mejora futura, pendiente de que el equipo de Infraestructura de Base de Datos habilite primero el soporte TLS en su lado; no se puede activar unilateralmente desde la aplicación.

### M7 — La base de datos de producción conserva el nombre `eir_test`

**Hallazgo:** por motivos históricos, la base de datos sobre la que corre hoy la aplicación sigue llamándose `eir_test`, lo que puede inducir a confusión sobre su carácter de producción.

**Decisión:** se resolverá como parte de la transición de aplicaciones ya planificada (migración a la base de datos definitiva `eir`), fuera del alcance de este ejercicio de hardening.

## Bajo

### B1 — El diálogo de confirmación al bloquear un IMEI no se mostraba

**Hallazgo:** un error de sintaxis en el código JavaScript de la pantalla de alta de bloqueo hacía que el diálogo de confirmación nunca llegara a mostrarse; el formulario se enviaba directamente. No era un problema de seguridad (el servidor sigue validando todo igual), solo se perdía una salvaguarda de experiencia de usuario para una acción irreversible.

**Mitigación:** corregido el error de sintaxis.

### B2 — Falta de comprobación de permiso explícita en el alta de solicitudes

**Hallazgo:** la pantalla final de alta de una solicitud dependía indirectamente de que otra pantalla anterior hubiera comprobado el permiso correcto, en vez de comprobarlo ella misma. No era explotable hoy, pero era una dependencia frágil ante futuros cambios.

**Mitigación:** añadida la comprobación de permiso explícita, igual que tienen el resto de pantallas sensibles del proyecto.

### B3 — Comparación del token de seguridad (CSRF) no resistente a ataques de temporización

**Hallazgo:** la comparación del token de protección contra CSRF usaba el operador de comparación estándar, que en teoría podría filtrar información por el tiempo de respuesta.

**Mitigación:** se reutilizó la misma función de comparación en tiempo constante ya usada para los hashes de contraseña.

### B4 — Escapado incompleto del nombre de un fichero en el listado de Ficheros

**Hallazgo:** el nombre de fichero se insertaba en un atributo de evento de un enlace quitando manualmente solo algunos caracteres especiales, sin cubrir todos los que podrían romper el HTML. El nombre de fichero no es hoy controlable por ningún usuario de la aplicación (lo genera un proceso externo), pero seguía siendo un punto débil.

**Mitigación:** el valor se mueve a un atributo de datos con la misma codificación HTML que ya protege el resto de la fila, en vez de insertarse directamente en el código del evento.

### B5 — Sin validación de formato al editar parámetros numéricos del sistema

**Hallazgo:** la pantalla de administración de parámetros permitía guardar cualquier texto en parámetros que la aplicación espera como números; un valor no numérico se ignoraba en silencio y se usaba el valor por defecto, sin avisar al administrador de que su cambio no había surtido efecto.

**Mitigación:** los parámetros numéricos conocidos ahora validan el formato antes de guardar, con un mensaje de error claro si no es válido.

## Nota técnica transversal

Durante este periodo se completó también la migración de toda la criptografía de la aplicación (hashes de contraseña, tokens de seguridad) de JScript a VBScript puro, al descubrirse una incompatibilidad real entre ambos motores de scripting en el servidor de producción (Windows Server 2025) que causaba fallos intermitentes del sitio. Esta migración, aunque motivada por un problema de estabilidad y no de seguridad, es la base técnica sobre la que se apoyan varias de las correcciones de este informe (M3 en particular).
