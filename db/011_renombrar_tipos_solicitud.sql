-- 011 - Renombrar los tipos de solicitud (Inclusion/Exclusion) - ToperBTR
-- ============================================================
-- Contexto (2026-10-29): peticion de negocio - el texto "Inclusion"/"Exclusion" (nombres
-- heredados del legacy Java) se sustituye por "Bloqueo Terminal"/"Desbloqueo Bloqueo
-- Terminal" en toda la aplicacion. Afecta al desplegable de alta de solicitud
-- (app/modules/solicitudes/pre_alta.asp), al listado y al detalle de solicitudes
-- (consulta.asp/detalle.asp) - los tres leen TIPO_SOLICITUD_REF.DESCRIPCION directamente,
-- asi que este UPDATE es el unico cambio necesario para que el texto nuevo aparezca en
-- los tres sitios a la vez. Los mensajes de error y titulos de pagina que tenian el texto
-- literal "Inclusion"/"Exclusion" escrito a mano en el codigo (alta.asp, alta_do.asp,
-- pre_alta_do.asp) se han actualizado aparte, directamente en el codigo fuente.
--
-- Se actualiza por CLAVE ('I'/'E'), no por DESCRIPCION ni por COD_TIPO_SOLICITUD: la
-- CLAVE es el identificador estable que usa el codigo (CodigoTipoSolicitudPorClave,
-- ver app/include/data/catalogos.asp) y no cambia con este script, asi que es la forma
-- mas segura de localizar la fila correcta sin depender del texto que se esta
-- reemplazando ni de un codigo numerico que podria variar entre entornos.
--
-- Confirmado en el entorno de pre (eir_test) el estado antes de este cambio:
--   cod_tipo_solicitud | descripcion | clave
--   1                  | Inclusion   | I
--   2                  | Exclusion   | E
--
-- UPDATE simple sobre 2 filas como mucho - no hace falta transaccion explicita, es
-- atomico por sentencia igual que el resto de UPDATE de este proyecto. Idempotente:
-- volver a ejecutarlo con los valores ya puestos no cambia nada ni da error.
--
-- IMPORTANTE: antes de aplicar en un entorno nuevo, confirmar que las filas con esas
-- claves existen y tienen el texto esperado:
--   SELECT COD_TIPO_SOLICITUD, DESCRIPCION, CLAVE FROM TIPO_SOLICITUD_REF WHERE CLAVE IN ('I', 'E');

USE eir_test;

UPDATE TIPO_SOLICITUD_REF SET DESCRIPCION = 'Bloqueo Terminal' WHERE CLAVE = 'I';
UPDATE TIPO_SOLICITUD_REF SET DESCRIPCION = 'Desbloqueo Terminal' WHERE CLAVE = 'E';
