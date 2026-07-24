# 8. Cambios de BBDD — ToperBTR

Estado: cerrado a nivel de lista y categorización. Se aplican **en vivo sobre la BD de producción `eir`** (no hay migración a una BD nueva — decisión confirmada 2026-07-23), con periodo de convivencia entre la app legacy y ToperBTR.

## Imprescindibles

| Cambio | Es seguro en convivencia | Notas |
|---|---|---|
| `usuarios.password` / `historico_password.password` → `varchar(255)` | Sí — ampliar una columna no rompe al legacy, que sigue escribiendo valores cortos (SHA1+Base64) | Aplicar primero, es la base de todo lo demás |
| Nueva columna `usuarios.algoritmo_password` | Sí — columna nueva, el legacy la ignora | Rellenar `SHA1B64` para todos los usuarios existentes al aplicar el cambio |
| `AUTO_INCREMENT` en: `usuarios`, `solicitud`, `cliente`, `lista_negra`, `lista_blanca`, `historico_bloqueo_directo`, `ficheros`, `incidencia`, `comentario_incidencia` | Sí — el legacy sigue insertando IDs explícitos sin problema; InnoDB ajusta el contador al alza automáticamente | Fijar el contador a `MAX(id)+1` real de cada tabla al aplicar el cambio; revisar antes si hay colisiones/huecos ya existentes por la condición de carrera histórica del `MAX+1` legacy |
| Fila `estado_usuario_ref (2, 'Bloqueado')` | Sí, aditivo | **Verificar primero** con `SELECT * FROM estado_usuario_ref` — el código legacy ya referencia `UsuarioVO.BLOQUEADO=2` activamente (solo la escritura estaba desactivada), es posible que la fila ya exista y no haga falta insertarla |
| Reutilizar claves ya existentes en `parametros_ref`: `NUM_FALLOS_LOGIN_USUARIO`, `NUM_OLD_PASSWORD`, `DURACION_PASSWORD`, `COD_OPER_HOME` | Sí | El legacy ya las lee — confirmar sus valores actuales (`SELECT * FROM parametros_ref`) antes de asumir ninguno; no crear claves duplicadas |
| Nueva fila `parametros_ref` (`NUM_MAX_INTENTOS_FALLIDOS`, valor por defecto 5) | Sí, aditivo | Nueva, para el bloqueo automático de ToperBTR (no existía como tal en el legacy activo) |

## Recomendables

| Cambio | Notas |
|---|---|
| Tabla nueva `auditoria_accion` (usuario, acción, entidad, fecha, IP, resultado) | Aditiva, sin relación con tablas legacy — soporta el requisito de auditoría mínima |

## Ya resueltos, sin cambio de BBDD necesario

| Punto | Resolución |
|---|---|
| Fichero "genérico" (`cod_fichero=0`) excluido en búsquedas | Resuelto por código: es solo un valor "ID no asignado", no un fichero especial. Con `AUTO_INCREMENT` los IDs reales nunca serán 0 — basta con seguir excluyendo `cod_fichero=0` en la consulta, sin necesidad de parametrizarlo |
| Código de negocio `incidencia.cod_incidencia` a 6 dígitos (antes 5, riesgo de overflow ya existente en el legacy) | No es cambio de BBDD — la columna ya es `varchar(15)`, cabe de sobra; es solo un cambio de lógica de aplicación (padding a 6 en vez de 5) |

## No recomendados

| Cambio | Motivo |
|---|---|
| Añadir FK real a `solicitud.cod_ind_solicitud_padre` | Podría romper inserciones de otros procesos (EIRControl) que no respeten el orden de creación |
| Eliminar `incidencia_test`, `solicitud_pruebas`, `suboperador_ref` | Confirmar con negocio/EIRControl antes de tocarlas — no hay evidencia de que estén libres de uso |

## Orden de aplicación recomendado

1. Todos los cambios "Imprescindibles" (son aditivos, cero riesgo para el legacy).
2. Verificar que la app legacy sigue funcionando igual tras aplicarlos.
3. Empezar a desplegar los módulos de ToperBTR según `07-plan-fases.md`.

## Pendiente de verificar contra la BD real (antes de aplicar los cambios)

- `SELECT * FROM estado_usuario_ref` — confirmar si ya existe la fila `Bloqueado`.
- `SELECT * FROM parametros_ref` — confirmar valores actuales de `NUM_FALLOS_LOGIN_USUARIO`, `NUM_OLD_PASSWORD`, `DURACION_PASSWORD`, `COD_OPER_HOME`.
- Revisar si ya hay colisiones/huecos de ID en las tablas afectadas por `AUTO_INCREMENT`, derivados de la condición de carrera histórica del patrón `MAX+1` del legacy.
