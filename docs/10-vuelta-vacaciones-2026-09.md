# Revisión post-vacaciones — septiembre 2026

Commit de referencia: `452aeae`  
Rama: `main`

---

## 1. Cambios manuales integrados (lo que había al volver)

### Nuevos módulos completos

| Módulo | Ficheros añadidos | Estado |
|---|---|---|
| **Bloqueos (E6)** | `bloqueos/bloqueo.asp`, `bloquear.asp`, `bloquear_do.asp`, `desbloquear_do.asp` | OK |
| **Admin (E7)** | `admin/parametros.asp`, `parametro_do.asp` | OK |
| **DB** | `db/006_bloqueos.sql`, `db/008_verificacion_predespliegue.sql` | OK |
| **Docs** | `docs/09-plan-pruebas.md` | OK |

### Mejoras funcionales destacadas

- **`auth.asp`** — rehash automático se activa también cuando el número de iteraciones PBKDF2 almacenado en el hash difiere del configurado en `Application("PBKDF2Iteraciones")`, no solo cuando el algoritmo es distinto. Corrección necesaria tras el cambio 50k→200 iteraciones.
- **`menu.asp`** — elimina la query a BD por cada página; ahora usa `CodigoAccionCacheado()` y `CodigoPerfilCacheado()` (datos en Application scope, cargados en `global.asa`).
- **`catalogos.asp`** — nuevas funciones: `CodigoAccionCacheado`, `CodigoPerfilCacheado`, `ListarTiposBloqueo`, `ListarRazones`, `CodigoRazon`.
- **`solicitudes_dal.asp`** — `BuscarSolicitudes` usa `CONCAT_WS` para devolver el cliente en una columna `CLIENTE`; el JOIN de estado vigente usa `MAX(FECHA_INICIO)` en lugar del `FECHA_FIN IS NULL` del legacy (más robusto si hay solapamientos).
- **`ficheros_dal.asp`** — amplía el reconocimiento del operador Yoigo/Grupo MASMOVIL/GMM para el filtro EIR.
- **`ficheros/descargar.asp`** — la ruta física se construye con `Server.MapPath("/ficherosBTR/")` + nombre de BD; elimina el path traversal residual que quedaba al concatenar la ruta almacenada.
- **`login_do.asp`** — `Server.ScriptTimeout = 360` para permitir que los usuarios con hashes legacy (50k iteraciones) puedan migrar en su primer login. **Eliminar este timeout cuando todos los usuarios hayan hecho login al menos una vez con el nuevo servidor.**
- **`Response.End` tras todos los `Response.Redirect`** — corrige el comportamiento de IIS Classic ASP donde la ejecución continuaba después del redirect.

### UI / Assets

- Logo PNG (`ToperBTR_Logo_peque_v3.png`) en pantalla de login y cabecera.
- Iconos SVG inline para "Cambiar contraseña" y "Desconectar" en la cabecera.
- CSS: `.form-edicion`, `.login-logo`, `.login-subtitulo`, `.cabecera-icono`.

### Hardening (web.config)

- `errorMode="DetailedLocalOnly"` (el original era `Detailed`, exponía stack traces a todos).
- Security headers: `X-Frame-Options: SAMEORIGIN`, `X-Content-Type-Options: nosniff`, `X-XSS-Protection`.
- `requestFiltering`: `allowDoubleEscaping=false`, límites URL/body.
- Error pages mapeadas: 401→sinpermiso, 403→sinpermiso, 404→noencontrado, 500→error.asp.
- `web.config` en `app/dal/` y `app/include/` para denegar acceso HTTP directo a los includes.

---

## 2. Bugs encontrados y corregidos en la revisión

| # | Fichero | Bug | Impacto | Corrección |
|---|---|---|---|---|
| 1 | `dal/bloqueos_dal.asp:46` | `Response.Write "No hay bloqueo<br>"` de debug en `ExcluirIMEI` | **Crítico** — inyectaba HTML en cualquier respuesta donde el IMEI no tuviera bloqueo activo (caso normal en Inclusiones) | Eliminado |
| 2 | `dal/solicitudes_dal.asp` | `BuscarSolicitudPadrePorImei` reescrita sin JOIN a CLIENTE; `pre_alta_do.asp` seguía leyendo `padre("NombreCliente")` etc. | **Crítico** — los campos de prefill del formulario de Exclusión quedaban vacíos | Restaurado JOIN con CLIENTE y nombres originales de claves del Dictionary |
| 3 | `include/layout/menu.asp` | Enlace `<a href="/conn_test.asp">Prueba conexion BTR</a>` fuera de `<li>`, fuera de bloque `<% %>`, expuesto en producción | **Medio** — HTML malformado + ruta de debug visible a usuarios | Eliminado |
| 4 | `modules/bloqueos/bloqueo.asp:10-12` | `CodigoPerfil(conn, "Administrador")` y `CodigoPerfil(conn, "Tramitación")` lanzaban 2 queries a BD en cada carga | **Menor** — inconsistente con la optimización hecha en menu.asp | Sustituido por `CodigoPerfilCacheado()` |

---

## 3. Deuda técnica / inconsistencias pendientes

### 3.1 `login_do.asp` — timeout de migración de hashes

```asp
' TEMPORAL: eliminar cuando todos los usuarios hayan hecho login con el nuevo servidor
Server.ScriptTimeout = 360
```

Fecha límite recomendada: verificar con el administrador de BD cuándo todos los hashes en `USUARIOS.PASSWORD_HASH` tengan 200 iteraciones (script de comprobación en `db/008_verificacion_predespliegue.sql`).

### 3.2 `solicitudes_dal.asp` — formato de `GenerarCodigoSolicitud` cambiado

El nuevo formato es `Y<YYYYMMDD><00000+id><clave>` (ej: `Y202609080001I`). El antiguo era `<clave><YYYYMMDD><000000+id>`. **Verificar con negocio si este cambio es compatible con los sistemas que consumen el código de solicitud** (ficheros EIR, confirmaciones de operadores).

### 3.3 `pre_alta_do.asp` — `Sol_PrefillBloqueado` forzado a `False`

```asp
'Session("Sol_PrefillBloqueado") = True   ← original
Session("Sol_PrefillBloqueado") = False   ← cambiado manualmente
```

Confirmar si `alta.asp` usa este flag para algo relevante en el formulario de Exclusión (mostrar/ocultar campo de bloqueo, deshabilitar campos, etc.).

### 3.4 `bloqueo.asp` — `RequierePerfil` con `Array()`

`RequierePerfil` recibe un array con los códigos de perfil cacheados. Si `global.asa` no ha cargado todavía los catálogos (primer arranque del pool de aplicación), los valores serán `-1` y se denegará el acceso a todos. Añadir una ruta de fallback en `global.asa` o en `RequierePerfil` para ese caso de arranque en frío.

### 3.5 Ficheros no versionados (actualmente `??` en git)

| Fichero/dir | Propósito | Acción recomendada |
|---|---|---|
| `app/conn_test.asp` | Script de prueba de conexión | Mover a `.gitignore` o eliminar antes de producción |
| `app/modules/admin/reset_hash_test.asp` | Script de reset manual de hash | Ídem — no debe estar accesible en producción |
| `FicherosBTR/` | Directorio de ficheros EIR | Añadir a `.gitignore` (datos operativos, no código) |
| `Modelo de código ASP...docx` | Documentación | Mover a `docs/` o `gitignore` |

---

## 4. Próximos pasos para continuar la refactorización

### Prioritario (antes de despliegue)

- [ ] **Eliminar `Server.ScriptTimeout = 360`** en `login_do.asp` cuando los hashes estén migrados.
- [ ] **Añadir `.gitignore`**: `FicherosBTR/`, `*.docx`, `conn_test.asp`, `reset_hash_test.asp`.
- [ ] **Verificar formato de `GenerarCodigoSolicitud`** con negocio/BA.
- [ ] **Confirmar `Sol_PrefillBloqueado = False`**: revisar si `alta.asp` usa este valor.

### Refactorización siguiente (E10 propuesto)

- [ ] **Revisar `global.asa`**: confirmar que carga todos los catálogos necesarios para `CodigoAccionCacheado` y `CodigoPerfilCacheado` (catálogos `ACCION_REF`, `PERFIL_REF`).
- [ ] **Completar bloquear.asp / bloquear_do.asp**: revisar si ya valida el IMEI contra `LISTA_BLANCA` (verificación Yoigo/GMM) antes de insertar en `LISTA_NEGRA`.
- [ ] **Módulo Informes**: el legacy tenía generación de ficheros EIR. Pendiente de análisis si está en scope.
- [ ] **Unificar indentación**: mezcla de espacios y tabs en los ficheros retocados manualmente (no funcional, pero dificulta el diff).
- [ ] **Test de humo sobre el plan de pruebas** (`docs/09-plan-pruebas.md`): ejecutar los casos del bloque A (login) y B (solicitudes) con la BD real.

### Deuda de código a limpiar

- [ ] `solicitudes_dal.asp` — SQL comentado en `GenerarCodigoSolicitud` (ya limpiado el de `BuscarSolicitudPadrePorImei` y `BuscarSolicitudes`, queda el bloque de una línea comentada en la función).
- [ ] `bloqueos_dal.asp` — indentación mixta en `ExcluirIMEI` (tabs vs espacios).
- [ ] `consulta.asp` GET branch — la query vacía devuelve `NOMBRE_CLIENTE`/`PRIMER_APELLIDO` pero el template renderiza la columna `CLIENTE` (sin impacto runtime porque `WHERE 1=0` nunca itera, pero es confuso).

---

## 5. Estado del repositorio

```
Branch: main
Commit: 452aeae
Entregables completados: 1–9
Archivos de aplicación: app/**  (≈ 50 ficheros .asp)
Scripts BD: db/001–008
Docs: docs/01–10
```
