-- Reaplicacion idempotente de los scripts 001-006 sobre `eir_test` - ToperBTR
-- ============================================================
-- Contexto (2026-09-14): conn_test.asp ("Diagnostico PKs sin AUTO_INCREMENT,
-- ronda 2") muestra que CLIENTE, HISTORICO_BLOQUEO_DIRECTO y LISTA_NEGRA siguen
-- sin AUTO_INCREMENT, y que HISTORICO_PASSWORD no tiene la columna
-- ALGORITMO_PASSWORD. Eso indica que esta instancia de `eir_test` no tiene
-- aplicado NINGUNO de los scripts 001-006 (probablemente se refresco desde un
-- dump de produccion despues de la ejecucion de julio).
--
-- Este script es el equivalente a ejecutar 001+002+003+004+006 en orden, pero
-- escrito para ser seguro de re-ejecutar tantas veces como haga falta sin
-- importar en que punto quedo la BD:
--   - ALTER ... MODIFY COLUMN (anchura, AUTO_INCREMENT) es idempotente por
--     naturaleza en MySQL: reaplicar la misma definicion no da error.
--   - ADD COLUMN no lo es (da "Duplicate column name" si ya existe), asi que
--     se protege con SQL dinamico que comprueba information_schema antes.
--   - Los INSERT centinela ya usaban WHERE NOT EXISTS en los scripts
--     originales; se mantiene igual aqui.
--   - CREATE TABLE usa IF NOT EXISTS igual que el script 001 original.
--
-- Los scripts 001-006 individuales NO se tocan (quedan como historico de lo
-- que se penso/aplico en cada entregable). Este script pasa a ser la forma
-- recomendada de dejar `eir_test` al dia a partir de ahora.

USE eir_test;

-- ============================================================
-- Helper: añadir columna solo si no existe ya
-- ============================================================
DROP PROCEDURE IF EXISTS _toperbtr_add_column_if_missing;
DELIMITER $$
CREATE PROCEDURE _toperbtr_add_column_if_missing(
    IN p_tabla VARCHAR(64),
    IN p_columna VARCHAR(64),
    IN p_ddl VARCHAR(500)
)
BEGIN
    IF NOT EXISTS (
        SELECT 1 FROM information_schema.COLUMNS
        WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = p_tabla AND COLUMN_NAME = p_columna
    ) THEN
        SET @ddl_sql = p_ddl;
        PREPARE stmt FROM @ddl_sql;
        EXECUTE stmt;
        DEALLOCATE PREPARE stmt;
    END IF;
END$$
DELIMITER ;

-- ============================================================
-- 001 - Base + Login + Autenticacion + Autorizacion
-- ============================================================

ALTER TABLE usuarios MODIFY COLUMN password VARCHAR(255) NOT NULL;
ALTER TABLE historico_password MODIFY COLUMN password VARCHAR(255) NOT NULL;

CALL _toperbtr_add_column_if_missing('usuarios', 'algoritmo_password',
    'ALTER TABLE usuarios ADD COLUMN algoritmo_password VARCHAR(20) NOT NULL DEFAULT ''SHA1B64''');
UPDATE usuarios SET algoritmo_password = 'SHA1B64' WHERE algoritmo_password = '';

-- VERIFICAR antes si hay colisiones/huecos por la condicion de carrera del MAX+1 legacy:
-- SELECT cod_indice_usuario, COUNT(*) FROM usuarios GROUP BY cod_indice_usuario HAVING COUNT(*) > 1;
INSERT INTO estado_usuario_ref (estado, descripcion)
SELECT 2, 'Bloqueado'
WHERE NOT EXISTS (SELECT 1 FROM estado_usuario_ref WHERE estado = 2);

INSERT INTO parametros_ref (clave, valor)
SELECT 'NUM_MAX_INTENTOS_FALLIDOS', '5'
WHERE NOT EXISTS (SELECT 1 FROM parametros_ref WHERE clave = 'NUM_MAX_INTENTOS_FALLIDOS');

ALTER TABLE usuarios MODIFY COLUMN cod_indice_usuario INT NOT NULL AUTO_INCREMENT;

CREATE TABLE IF NOT EXISTS auditoria_accion (
  id INT NOT NULL AUTO_INCREMENT,
  cod_indice_usuario INT NULL,
  accion VARCHAR(50) NOT NULL,
  entidad VARCHAR(50) NOT NULL,
  entidad_id VARCHAR(50) NULL,
  fecha DATETIME NOT NULL,
  ip VARCHAR(45) NULL,
  resultado VARCHAR(20) NOT NULL,
  detalle VARCHAR(500) NULL,
  PRIMARY KEY (id),
  KEY idx_auditoria_usuario (cod_indice_usuario),
  KEY idx_auditoria_fecha (fecha)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ============================================================
-- 002 - Usuarios (alta/baja/desbloqueo/cambio password)
-- ============================================================

CALL _toperbtr_add_column_if_missing('historico_password', 'algoritmo_password',
    'ALTER TABLE historico_password ADD COLUMN algoritmo_password VARCHAR(20) NOT NULL DEFAULT ''SHA1B64''');
UPDATE historico_password SET algoritmo_password = 'SHA1B64' WHERE algoritmo_password = '';

-- ============================================================
-- 003 - Solicitudes
-- ============================================================

ALTER TABLE solicitud MODIFY COLUMN cod_indice_sol INT NOT NULL AUTO_INCREMENT;
ALTER TABLE cliente MODIFY COLUMN cod_cliente INT NOT NULL AUTO_INCREMENT;
ALTER TABLE historico_bloqueo_directo MODIFY COLUMN cod_hist_bloq_dir INT NOT NULL AUTO_INCREMENT;

-- VERIFICAR primero si ya existe: SELECT * FROM ficheros WHERE cod_fichero = 0;
INSERT INTO ficheros (cod_fichero, nombre, fecha_creacion, procesado)
SELECT 0, 'GENERICO_TOPERBTR', NOW(), 1
WHERE NOT EXISTS (SELECT 1 FROM ficheros WHERE cod_fichero = 0);

INSERT INTO parametros_ref (clave, valor)
SELECT 'COD_FICHERO_GENERICO', '0'
WHERE NOT EXISTS (SELECT 1 FROM parametros_ref WHERE clave = 'COD_FICHERO_GENERICO');

-- ============================================================
-- 004 - Incidencias
-- ============================================================

ALTER TABLE comentario_incidencia MODIFY COLUMN cod_indice_comentario INT NOT NULL AUTO_INCREMENT;

-- ============================================================
-- 006 - Bloqueos
-- ============================================================

-- VERIFICAR antes si hay huecos/colisiones: SELECT MAX(cod_lista_negra) FROM lista_negra;
ALTER TABLE lista_negra MODIFY COLUMN cod_lista_negra INT NOT NULL AUTO_INCREMENT;

-- ============================================================
-- Limpieza del helper
-- ============================================================
DROP PROCEDURE IF EXISTS _toperbtr_add_column_if_missing;

-- ============================================================
-- FIN. A continuacion:
--   1) Volver a ejecutar db/008_verificacion_predespliegue.sql -> todo OK.
--   2) Volver a ejecutar app/conn_test.asp -> las 3 tablas deben salir en verde
--      y HISTORICO_PASSWORD debe seguir marcada como "PK compuesta" (correcto,
--      no es un problema, es una clave natural).
-- ============================================================
