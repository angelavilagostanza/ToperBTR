-- Entregable 3 (Solicitudes) - ToperBTR
-- Se aplica en vivo sobre la BD copia de produccion `eir_test` (ver docs/08-cambios-bbdd.md).


USE eir_test;

-- 1) AUTO_INCREMENT (MySQL fija el contador automaticamente al MAX(id) existente al
--    hacer el ALTER). historico_bloqueo_directo se necesita ya en este entregable porque
--    una solicitud de Exclusion puede disparar una exclusion de bloqueo (ver dal/bloqueos_dal.asp).
ALTER TABLE solicitud MODIFY COLUMN cod_indice_sol INT NOT NULL AUTO_INCREMENT;
ALTER TABLE cliente MODIFY COLUMN cod_cliente INT NOT NULL AUTO_INCREMENT;
ALTER TABLE historico_bloqueo_directo MODIFY COLUMN cod_hist_bloq_dir INT NOT NULL AUTO_INCREMENT;

-- 2) Fichero centinela para solicitudes creadas desde la web.
--    SOLICITUD.COD_FICHERO es NOT NULL con FK a FICHEROS, pero una solicitud creada
--    desde ToperBTR no viene de ningun fichero de intercambio real - hace falta una fila
--    que represente "no aplica / creada por la web", igual que ya hacia el legacy con
--    COD_FICHERO=0 (Constantes.FICHERO_GENERICO, confirmado en el codigo Java real).
--    VERIFICAR primero si ya existe: SELECT * FROM ficheros WHERE cod_fichero = 0;
INSERT INTO ficheros (cod_fichero, nombre, fecha_creacion, procesado)
SELECT 0, 'GENERICO_TOPERBTR', NOW(), 1
WHERE NOT EXISTS (SELECT 1 FROM ficheros WHERE cod_fichero = 0);

INSERT INTO parametros_ref (clave, valor)
SELECT 'COD_FICHERO_GENERICO', '0'
WHERE NOT EXISTS (SELECT 1 FROM parametros_ref WHERE clave = 'COD_FICHERO_GENERICO');
