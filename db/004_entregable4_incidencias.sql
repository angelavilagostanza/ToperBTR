-- Entregable 4 (Incidencias) - ToperBTR
-- Se aplica en vivo sobre la BD copia de produccion `eir_test` (ver docs/08-cambios-bbdd.md).
--
-- Este entregable NO crea incidencias desde la web (se confirmo que "alta de incidencia"
-- en el legacy es un proceso automatico del sistema/EIRControl al procesar ficheros, no
-- una accion manual de usuario; el unico "alta" que ofrece la web es anadir un
-- comentario). Por eso solo hace falta AUTO_INCREMENT en COMENTARIO_INCIDENCIA, no en
-- INCIDENCIA.

ALTER TABLE comentario_incidencia MODIFY COLUMN cod_indice_comentario INT NOT NULL AUTO_INCREMENT;
