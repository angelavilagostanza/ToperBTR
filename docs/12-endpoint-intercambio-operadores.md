# 12. Endpoint de intercambio para operadores externos

Documento de referencia para el administrador de IIS/infraestructura. Describe qué hace
el nuevo endpoint `app/externos/intercambio/descargar.asp` y qué hay que configurar en
el servidor para que funcione — nada de esto se puede hacer desde el código del
repositorio, requiere acceso a IIS Manager / `applicationHost.config`.

## Qué reemplaza

En el legacy (`EIRWeb`), los sistemas de otros operadores descargaban el fichero diario
de intercambio EIR (altas/bajas de IMEI en lista negra) llamando a:

```
https://www.btr.yoigo.es/EIRWeb/ficheros.do?cmd=publicarFichero&action=abrir&nombreFichero=Y20201213
```

Esa URL pública exigía **certificado de cliente (mTLS)** antes de reenviar la petición a
un endpoint interno (`http://smadavbtr.yoigo.inet:8180/EIRWeb/ficheros.do?...`) que
ejecutaba la lógica real sin volver a comprobar nada — confiaba en que el salto anterior
ya había autenticado la llamada.

`app/externos/intercambio/descargar.asp` reproduce la mitad "interna" de ese modelo:
recibe `GET ?fichero=<nombre>`, valida que el nombre tiene la forma de un fichero de
intercambio real (1 letra de operador + 8 dígitos de fecha, p.ej. `Y20260802`), lo busca
en la tabla `FICHEROS` y lo devuelve en binario. **No comprueba el certificado de
cliente por su cuenta** — eso es responsabilidad de IIS, en esta misma ruta, antes de
que la petición llegue al `.asp`.

La URL pública definitiva **se coordinará con los operadores por separado** (no hace
falta replicar la ruta `.do` legacy) — este documento asume que ya se ha decidido.

## Checklist a aplicar en el servidor

1. **Vdir/ruta**: publicar `app/externos/intercambio/` en la URL que se acuerde con los
   operadores.
2. **Exigir SSL + certificado de cliente** en esa ruta concreta (no en todo el sitio):
   - Vía IIS Manager (SSL Settings → Require SSL + Client certificates: Require), o
   - Vía `web.config` con `<security><access sslFlags="Ssl, SslNegotiateCert,
     SslRequireCert" /></security>` — **solo posible si esa sección no está bloqueada a
     nivel de servidor**; si lo está, desplegar esa configuración produce un error
     `500.19` en cada petición. Confirmar con el administrador cuál de las dos vías se
     usa antes de tocar el `web.config`.
3. **CA(s) de confianza**: configurar qué entidad(es) certificadora(s) se aceptan para
   los certificados de cliente de los operadores en esa ruta.
4. **Autenticación anónima**: confirmar que sigue **permitida** en esta ruta. El
   `web.config` de ToperBTR no define Forms/Windows Authentication en ningún sitio, y el
   propio script no comprueba sesión — si IIS tiene la autenticación anónima
   deshabilitada a nivel de sitio/aplicación (fuera de lo que ve este repositorio), la
   petición se rechazaría antes de llegar al `.asp` aunque el certificado fuera válido.
5. **Verificar el override de `httpErrors`** de `app/externos/intercambio/web.config`
   (`existingResponse="PassThrough"`) realmente toma efecto — si esa sección también
   está bloqueada a nivel de servidor, dará `500.19` en vez de aplicarse en silencio.
   Probar después de desplegar:
   - `?fichero=` (vacío) o `?fichero=XX` (formato inválido) → debe devolver **400** en
     texto plano, no una redirección a una página HTML.
   - `?fichero=Y20991231` (formato válido, no existe) → debe devolver **404** en texto
     plano.
   - Un nombre real existente en el servidor → **200** con el contenido binario y la
     cabecera `Content-Disposition: attachment`.
6. **Límite de tamaño de petición**: `app/web.config` (raíz) ya limita
   `maxAllowedContentLength` a 4 MB, pero es sobre el *cuerpo de la petición* — esta
   llamada es un `GET` sin cuerpo, así que no aplica y no limita el tamaño del fichero
   descargado. No hace falta tocarlo.

## Qué NO hace este endpoint (a propósito)

- No valida el certificado de cliente en código (`Request.ClientCertificate`) — confía
  por completo en que IIS ya lo hizo antes de ejecutar el script, igual que el legacy.
- No acepta cualquier fichero de la tabla `FICHEROS` — solo los que tienen la forma de
  fichero de intercambio diario real. Los ficheros de tipo "Consolidación" o el feed
  interno "EIR" de Yoigo no son accesibles por esta vía.
- No usa sesión de navegador ni los mensajes de error HTML del resto de la aplicación
  (`/errors/error.asp`) — todas las respuestas de error son un código HTTP + texto
  plano, pensadas para un cliente automático, no para una persona con navegador.

## Auditoría

Cada intento (válido o rechazado) queda registrado en `auditoria_accion` con
`ACCION = 'FICHERO_DESCARGA_INTERCAMBIO'` (distinta de `FICHERO_DESCARGA`, que es la del
endpoint de descarga para administradores autenticados) — incluye la IP de origen
(capturada automáticamente) y el nombre de fichero solicitado, incluso cuando se
rechaza por formato inválido o no se encuentra. Es la única traza de quién usa este
endpoint, dado que no requiere login.
