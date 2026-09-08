-- Entregable 6 (Bloqueos) — ToperBTR
-- Se aplica en vivo sobre la BD de produccion `eir` (o `eir_test` en QA).
-- Ver docs/08-cambios-bbdd.md para el orden de aplicacion y notas de convivencia.

USE eir_test;

-- AUTO_INCREMENT en LISTA_NEGRA.
-- Los entregables anteriores ya aplicaron AUTO_INCREMENT en:
--   usuarios (002), solicitud/cliente/historico_bloqueo_directo (003),
--   incidencia/comentario_incidencia (004), ficheros (005).
-- LISTA_NEGRA es la unica tabla relevante para Bloqueos que aun usa el patron MAX+1 legacy.
-- LISTA_BLANCA la gestiona EIRControl (fuera de alcance de ToperBTR), se deja para un
-- script independiente coordinado con ese equipo.
-- MySQL fija el contador AUTO_INCREMENT al MAX(COD_LISTA_NEGRA)+1 existente en el momento
-- del ALTER, por lo que no hay riesgo de colision con filas ya existentes.
-- VERIFICAR antes si hay huecos/colisiones: SELECT MAX(COD_LISTA_NEGRA) FROM lista_negra;
ALTER TABLE lista_negra MODIFY COLUMN cod_lista_negra INT NOT NULL AUTO_INCREMENT;
