# 17. Arquitectura del nuevo webservice de intercambio para operadores (BTR-WS)

## Motivación: por qué no seguir con el endpoint ASP clásico actual

`app/externos/intercambio/descargar.asp` (ver [`docs/12-endpoint-intercambio-operadores.md`](12-endpoint-intercambio-operadores.md)) funciona, pero vive **dentro del mismo sitio, mismo pool de aplicación y mismo árbol de carpetas que el resto de ToperBTR**. Para un endpoint con exposición a internet (aunque sea detrás de un middleware con mTLS) eso tiene dos problemas:

1. **Blast radius**: un fallo, una sobrecarga o una vulnerabilidad futura en ese endpoint comparte proceso (`w3wp.exe`) y pool con la aplicación de uso interno. No está aislado.
2. **Tecnología**: ASP clásico no tiene hoy herramientas modernas de seguridad "de fábrica" para un servicio máquina-a-máquina expuesto (rate limiting, validación tipada, logging estructurado) — hay que construirlas a mano cada vez.

Decisión: **sacarlo a un webservice nuevo, completamente aislado**, con su propia URL pública (`btrws.yoigo.com`), detrás del middleware existente que ya autentica con certificado de cliente, y que entrega la petición ya autenticada a una **aplicación IIS nueva e independiente** en nuestro servidor.

## Topología

```
Internet
   |
   v
btrws.yoigo.com  (DNS publico)
   |
   v
Middleware / frontal existente  <-- exige y valida certificado de cliente (mTLS)
   |                                  (fuera del alcance de este repo, lo gestiona
   |                                   el equipo de red/infraestructura)
   v
Servidor IIS (el mismo servidor Windows Server 2025 que aloja ToperBTR)
   |
   +-- Sitio/Aplicacion "ToperBTR"      (pool propio, sin tocar)
   |
   +-- Sitio/Aplicacion "BTR-WS" nueva  (pool propio, independiente)
          |
          +-- Valida IP de origen contra tabla en BD (defensa adicional)
          +-- Valida formato de fichero + operador existente
          +-- Audita cada peticion (aceptada o rechazada)
          +-- Devuelve el fichero de intercambio
```

**Aislamiento real, no solo lógico**: pool de aplicación propio, identidad de pool propia (mínimo privilegio), carpeta física separada, sin código ni configuración compartida con ToperBTR. Si BTR-WS cae, se satura o se ve comprometido, ToperBTR no se entera.

## Defensa en profundidad (no confiar solo en el middleware)

Aunque el middleware ya exige certificado de cliente antes de dejar pasar la petición, el webservice nuevo no confía ciegamente en eso — añade sus propias capas, igual que ya hacía el legacy Java en su salto interno, pero reforzado:

1. **mTLS en el middleware** (capa externa, ya existente, fuera de nuestro control).
2. **Lista blanca de IP de origen contra una tabla en BD** (nueva, pedida explícitamente) — ver más abajo.
3. **Validación estricta de entrada**: mismo patrón que hoy (1 letra de operador + 8 dígitos de fecha) y comprobación de que el operador existe.
4. **Auditoría completa**: toda petición, aceptada o rechazada, queda registrada con IP de origen, fichero solicitado y resultado.
5. **Principio de mínimo privilegio**: el pool de aplicación corre con una identidad que solo puede leer lo estrictamente necesario (BD + carpeta de ficheros de intercambio), nada más.

**Pregunta abierta bloqueante para el día 1** (ver plan de acción): si el middleware reenvía la petición ya "traducida", la IP que verá IIS puede ser la del propio middleware, no la del operador real. Hay que confirmar con el equipo de red si existe una cabecera con la IP original (`X-Forwarded-For` o equivalente) — si no la hay, la capa 2 de arriba hay que rediseñarla (validar la IP del middleware, no la del operador, o apoyarse solo en mTLS).

## Tecnología elegida: ASP.NET Core 8 (LTS) Web API

Pensando en seguridad como factor principal y en la restricción de "sin terceros ni software adicional":

- **100% Microsoft**, parte del SDK/Runtime de .NET, gratuito, sin licencia ni proveedor externo.
- Se instala con el **ASP.NET Core Hosting Bundle** oficial de Microsoft (incluye el módulo IIS `AspNetCoreModuleV2`) — es runtime de Microsoft, no "software de terceros".
- Se ejecuta **in-process dentro de IIS** (`w3wp.exe`, modelo `InProcess`), en su propio pool — buen rendimiento y aislamiento real a la vez.
- Trae de fábrica, sin añadir ningún paquete externo: validación de modelos, *rate limiting* (`Microsoft.AspNetCore.RateLimiting`, incluido en el framework desde .NET 7), logging estructurado a Event Log o fichero, y acceso programático al certificado de cliente (`HttpContext.Connection.ClientCertificate`) por si se quiere una segunda comprobación además del middleware.
- **Acceso a la BD MySQL compartida (`eir_test`) vía `System.Data.Odbc`**, que es parte de la librería base de .NET — reutilizando el **mismo driver ODBC de MySQL ya instalado en el servidor** para ToperBTR. Esto evita instalar ningún conector de terceros (como el paquete NuGet de MySQL) solo para esto.
- IIS sigue siendo el único punto de escucha de internet — no hace falta exponer Kestrel directamente, coherente con la infraestructura actual.

## Estructura del proyecto

Repositorio **nuevo y separado** de `V:\ToperBTR` (nombre propuesto: `BTRWS`) — stack distinto (.NET vs VBScript), ciclo de despliegue independiente, y así el aislamiento de "blast radius" empieza ya en el control de versiones y no solo en IIS.

```
BTRWS/
  BtrWs.sln
  src/
    BtrWs.Api/
      BtrWs.Api.csproj
      Program.cs                    <- pipeline minimal API, middleware de orden:
                                        rate limiting -> IP allow-list -> endpoints
      appsettings.json              <- configuracion sin secretos
      appsettings.Production.json
      Endpoints/
        IntercambioEndpoints.cs     <- GET /intercambio/descargar
      Services/
        FicheroIntercambioService.cs
        IpAllowListService.cs
        AuditoriaService.cs
      Data/
        OdbcConnectionFactory.cs    <- misma cadena/DSN ODBC que ToperBTR
      web.config                    <- generado por "dotnet publish", con
                                        <aspNetCore hostingModel="InProcess">
  docs/
    (referencia a este documento y al plan de accion)
```

## Configuración en IIS

- **Pool de aplicación nuevo**: `BTRWS_AppPool`.
  - .NET CLR version: **"No Managed Code"** (obligatorio para ASP.NET Core).
  - Identidad: cuenta de pool dedicada, mínimo privilegio — sin compartir identidad con el pool de ToperBTR.
  - Start Mode: `AlwaysRunning` + Idle Timeout = 0, para evitar arranque en frío ante tráfico intermitente de operadores.
  - Reciclado: deshabilitado por horario, o solo en ventana de bajo tráfico (nunca mientras un operador puede estar descargando).
- **Sitio o Aplicación IIS nueva**, con *binding* propio (host header/puerto que acuerde el equipo de red para recibir el tráfico ya autenticado del middleware) — **recomendado: sitio nuevo**, no una subcarpeta del sitio de ToperBTR, para aislar también a nivel de *binding*.
- Instalar el **ASP.NET Core Hosting Bundle (.NET 8)** si el servidor no lo tiene ya — requiere `iisreset` tras la instalación.
- `httpErrors`: igual que se tuvo que resolver para el endpoint `.asp` actual (`existingResponse="PassThrough"`), confirmar que el `web.config` generado por `dotnet publish` no hereda la redirección HTML de IIS — ASP.NET Core ya devuelve sus propias respuestas (texto plano / `application/problem+json`) y no debe pasar por `/errors/error.asp` bajo ningún concepto, porque esta aplicación nueva ni siquiera tiene esa carpeta.
- Logging: Windows Event Log (proveedor incluido en .NET, sin software adicional) + fichero con rotación en una carpeta dedicada fuera del árbol de la aplicación.

## Base de datos: tabla de IPs permitidas

Tabla nueva en la **misma BD MySQL compartida pre/PRO** (`eir_test` en `172.30.170.27` — mismo dato de infraestructura que ya aplica a ToperBTR), para que más adelante sea sencillo gestionarla desde una pantalla de administración en ToperBTR (`toperbtr.masmovil.com`), que ya tiene esa misma conexión.

```sql
CREATE TABLE ws_ip_permitida (
  cod_ip_permitida   INT NOT NULL AUTO_INCREMENT,
  ip_o_rango         VARCHAR(43) NOT NULL,   -- IPv4/IPv6 exacta; deja hueco para CIDR futuro
  descripcion        VARCHAR(200) NULL,      -- p.ej. "Operador X - salida fija oficina"
  activo              TINYINT(1) NOT NULL DEFAULT 1,
  fecha_alta         DATETIME NOT NULL,
  cod_usuario_alta   INT NULL,               -- referencia informativa a usuarios.cod_indice_usuario,
  fecha_baja         DATETIME NULL,          -- sin FK: BTR-WS queda desacoplado del esquema de ToperBTR
  cod_usuario_baja   INT NULL,
  PRIMARY KEY (cod_ip_permitida),
  KEY idx_ip_activo (ip_o_rango, activo)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;
```

Decisiones de diseño:
- **Sin FK a `usuarios`**: BTR-WS no depende del esquema de ToperBTR evolucionando con el tiempo; las columnas `cod_usuario_alta`/`baja` son solo informativas para cuando exista la pantalla de administración.
- **Coincidencia exacta primero, no CIDR**: los operadores suelen llamar desde un número pequeño de IPs fijas; `ip_o_rango VARCHAR(43)` deja margen para IPv6 y para añadir soporte de rangos CIDR más adelante sin cambiar el esquema.
- Migración creada en `db/014_whitelist_ip_webservice.sql`, lista para aplicar en pre (recordar: pre y PRO comparten la misma BD física).

## Convivencia y retirada del endpoint actual

El `.asp` actual (`app/externos/intercambio/descargar.asp`) **no se retira de inmediato**. Plan:
1. BTR-WS queda operativo y probado.
2. Los operadores migran su integración a la nueva URL (`btrws.yoigo.com`), coordinado aparte.
3. Periodo de coexistencia con ambos endpoints activos.
4. Una vez confirmado que ningún operador usa ya la ruta antigua, se retira en un commit dedicado (código + *binding* de IIS), documentado en `docs/`.

## Preguntas abiertas (bloqueantes para el diseño fino, no para empezar)

1. ¿El middleware reenvía la IP real del operador (p.ej. cabecera `X-Forwarded-For`) o solo la suya propia? Condiciona cómo funciona de verdad la lista blanca de IP.
2. ¿El tramo middleware → IIS va cifrado (HTTPS interno) o es tráfico interno sin TLS? Afecta a si la aplicación nueva necesita certificado de servidor propio.
3. ¿Sitio IIS nuevo con *binding* propio, o aplicación dentro de un sitio ya existente pero en pool distinto? (Recomendación de este documento: sitio nuevo.)
4. ¿Está ya instalado el ASP.NET Core Hosting Bundle en el servidor, o hay que instalarlo?
5. Nombre definitivo del repositorio y de la carpeta de despliegue en el servidor.

Ver [`docs/18-plan-accion-webservice-intercambio.md`](18-plan-accion-webservice-intercambio.md) para cómo y cuándo resolver cada una.
