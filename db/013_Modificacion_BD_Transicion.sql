-- 013 - Consolidado de cambios de BD pendientes para la transicion - ToperBTR
-- ============================================================
-- PARA REVISAR CON EL EQUIPO DE INFRAESTRUCTURA DE BASE DE DATOS, antes de aplicar
-- nada contra la base de datos real de produccion (`eir`). Este fichero junta en un
-- solo sitio TODOS los cambios de esquema/datos que ToperBTR necesita y que todavia
-- no se han aplicado de forma consolidada, para que Infraestructura pueda valorar el
-- impacto de cada uno de una vez.
--
-- Contenido: las partes de db/009_reaplicacion_idempotente.sql (que ya consolidaba
-- 001+002+003+004+006) MAS los cambios nuevos de los scripts 010, 011 y 012, que
-- 009 no incluye por ser posteriores. Los ficheros individuales 001-012 NO se tocan
-- (quedan como historico de cada entregable); este fichero es la forma recomendada
-- de aplicar todo de una vez en un entorno nuevo.
--
-- Todos los cambios son ADITIVOS (nunca se borra ni se modifica una fila/columna ya
-- usada por el legacy de forma incompatible) y pensados para la convivencia con la
-- app legacy Java sobre la misma base de datos durante el periodo de transicion.
--
-- Idempotente: se puede volver a ejecutar sin que de error si una parte ya se aplico
-- antes (ALTER...MODIFY COLUMN es idempotente por naturaleza en MySQL; ADD COLUMN e
-- INSERT van protegidos con comprobacion previa).
--
-- IMPORTANTE para Infraestructura:
--   - Hacer backup completo de la base de datos antes de aplicar.
--   - Ninguna sentencia de este script hace DROP ni DELETE.
--   - Las columnas que se ensanchan (password, cod_usuario) solo se hacen MAS ANCHAS,
--     nunca mas estrechas - no puede truncar ningun dato existente.
--   - Los AUTO_INCREMENT que se activan fijan el contador automaticamente al valor
--     MAX(id) actual de cada tabla - no colisiona con las filas ya existentes.
--   - Revisar las dos consultas de verificacion que se piden ANTES de aplicar (sub-
--     seccion 1, mas abajo) - si devuelven filas, hay que resolverlo con el equipo de
--     aplicacion antes de continuar.
--
-- ============================================================
-- PASO 0 - VERIFICACIONES PREVIAS (solo lectura, ejecutar y revisar el resultado
--          ANTES de aplicar el resto del script)
-- ============================================================

-- 0.1) Duplicados en usuarios.cod_indice_usuario (condicion de carrera del MAX+1 del
--      legacy) - si esto devuelve alguna fila, NO continuar sin resolverlo primero,
--      el ALTER...AUTO_INCREMENT del Bloque 1 fallaria o dejaria datos inconsistentes.
SELECT cod_indice_usuario, COUNT(*) AS repeticiones
FROM usuarios
GROUP BY cod_indice_usuario
HAVING COUNT(*) > 1;

-- 0.2) Huecos/colisiones en lista_negra.cod_lista_negra, mismo motivo que 0.1.
SELECT cod_lista_negra, COUNT(*) AS repeticiones
FROM lista_negra
GROUP BY cod_lista_negra
HAVING COUNT(*) > 1;

-- 0.3) Confirmar el estado real de tipo_solicitud_ref antes del renombrado del Bloque 4
--      (se espera 'Inclusion'/clave 'I' y 'Exclusion'/clave 'E').
SELECT cod_tipo_solicitud, descripcion, clave FROM tipo_solicitud_ref WHERE clave IN ('I', 'E');

-- ============================================================
-- BLOQUE 1 - Entregables 1-4 y 6 (equivalente a db/009_reaplicacion_idempotente.sql)
-- ============================================================

-- Helper: añadir columna solo si no existe ya (se elimina al final del bloque).
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

-- --- 1.1) Hashing de contraseñas (PBKDF2SHA256 conviviendo con legacy SHA1B64) ---
-- Ensancha la columna para el nuevo formato autodescriptivo (algoritmo$iteraciones$
-- salt$hash); el legacy sigue escribiendo valores cortos sin ningun problema en una
-- columna mas ancha.
ALTER TABLE usuarios MODIFY COLUMN password VARCHAR(255) NOT NULL;
ALTER TABLE historico_password MODIFY COLUMN password VARCHAR(255) NOT NULL;

-- Columna discriminadora de algoritmo por fila, para que legacy y ToperBTR convivan
-- sin ambiguedad sobre con que algoritmo verificar cada hash.
CALL _toperbtr_add_column_if_missing('usuarios', 'algoritmo_password',
    'ALTER TABLE usuarios ADD COLUMN algoritmo_password VARCHAR(20) NOT NULL DEFAULT ''SHA1B64''');
UPDATE usuarios SET algoritmo_password = 'SHA1B64' WHERE algoritmo_password = '';

CALL _toperbtr_add_column_if_missing('historico_password', 'algoritmo_password',
    'ALTER TABLE historico_password ADD COLUMN algoritmo_password VARCHAR(20) NOT NULL DEFAULT ''SHA1B64''');
UPDATE historico_password SET algoritmo_password = 'SHA1B64' WHERE algoritmo_password = '';

-- --- 1.2) Catalogos y parametros nuevos que usa ToperBTR ---
-- El legacy ya referencia activamente UsuarioVO.BLOQUEADO=2; puede que esta fila ya
-- exista en produccion real (confirmarlo con el PASO 0 si hace falta).
INSERT INTO estado_usuario_ref (estado, descripcion)
SELECT 2, 'Bloqueado'
WHERE NOT EXISTS (SELECT 1 FROM estado_usuario_ref WHERE estado = 2);

-- Umbral de bloqueo por intentos fallidos de login (hallazgo de seguridad A1/A2,
-- ver docs/15-auditoria-seguridad-oct2026.md).
INSERT INTO parametros_ref (clave, valor)
SELECT 'NUM_MAX_INTENTOS_FALLIDOS', '5'
WHERE NOT EXISTS (SELECT 1 FROM parametros_ref WHERE clave = 'NUM_MAX_INTENTOS_FALLIDOS');

-- --- 1.3) AUTO_INCREMENT real (sustituye el patron MAX+1 del legacy, origen de
--     condiciones de carrera reales confirmadas en el codigo Java) ---
-- MySQL fija el contador automaticamente al MAX(id) existente en el momento del
-- ALTER - no colisiona con las filas ya existentes. Ver verificaciones del PASO 0.
ALTER TABLE usuarios MODIFY COLUMN cod_indice_usuario INT NOT NULL AUTO_INCREMENT;
ALTER TABLE solicitud MODIFY COLUMN cod_indice_sol INT NOT NULL AUTO_INCREMENT;
ALTER TABLE cliente MODIFY COLUMN cod_cliente INT NOT NULL AUTO_INCREMENT;
ALTER TABLE historico_bloqueo_directo MODIFY COLUMN cod_hist_bloq_dir INT NOT NULL AUTO_INCREMENT;
ALTER TABLE comentario_incidencia MODIFY COLUMN cod_indice_comentario INT NOT NULL AUTO_INCREMENT;
ALTER TABLE lista_negra MODIFY COLUMN cod_lista_negra INT NOT NULL AUTO_INCREMENT;

-- --- 1.4) Fichero y parametro centinela para solicitudes creadas desde la web ---
-- SOLICITUD.COD_FICHERO es NOT NULL con FK a FICHEROS, pero una solicitud creada
-- desde ToperBTR no viene de ningun fichero de intercambio real - hace falta una
-- fila que represente "no aplica / creada por la web", igual que ya hacia el legacy
-- con COD_FICHERO=0 (Constantes.FICHERO_GENERICO, confirmado en el codigo Java real).
INSERT INTO ficheros (cod_fichero, nombre, fecha_creacion, procesado)
SELECT 0, 'GENERICO_TOPERBTR', NOW(), 1
WHERE NOT EXISTS (SELECT 1 FROM ficheros WHERE cod_fichero = 0);

INSERT INTO parametros_ref (clave, valor)
SELECT 'COD_FICHERO_GENERICO', '0'
WHERE NOT EXISTS (SELECT 1 FROM parametros_ref WHERE clave = 'COD_FICHERO_GENERICO');

-- --- 1.5) Tabla nueva de auditoria de acciones (sin relacion con tablas del legacy) ---
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

DROP PROCEDURE IF EXISTS _toperbtr_add_column_if_missing;

-- ============================================================
-- BLOQUE 2 - Ampliar el login de usuario a 50 caracteres (db/010)
-- ============================================================
-- Decision de negocio: los usuarios nuevos se dan de alta de forma nominal y el
-- login pasa a ser la cuenta de correo electronico, que no cabe en el limite legacy
-- de 15 caracteres. ALTER...MODIFY COLUMN es idempotente, seguro de reaplicar.
ALTER TABLE usuarios MODIFY COLUMN cod_usuario VARCHAR(50) NOT NULL;

-- ============================================================
-- BLOQUE 3 - Enlace externo en Solicitud (db/012)
-- ============================================================
-- Nuevo campo para guardar un enlace a otra aplicacion (Jira, CRM, etc.) asociado a
-- la solicitud. Se rellena solo al dar de alta la solicitud; no hay pantalla de
-- edicion de solicitudes en ToperBTR, no se puede modificar despues de creada.
-- NULL por defecto: dato opcional, las solicitudes ya existentes no tienen nada que
-- migrar aqui. VARCHAR(500): mismo tamaño que ya usa el campo Comentario en otros
-- modulos del proyecto, de sobra para una URL normal con margen para enlaces largos.
SET @existe_columna_enlace = (
    SELECT COUNT(*) FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'solicitud' AND COLUMN_NAME = 'enlace_externo'
);

SET @sql_alter_enlace = IF(@existe_columna_enlace = 0,
    'ALTER TABLE solicitud ADD COLUMN enlace_externo VARCHAR(500) NULL',
    'SELECT 1'
);

PREPARE stmt FROM @sql_alter_enlace;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

-- ============================================================
-- BLOQUE 4 - Renombrar tipos de solicitud Inclusion/Exclusion (db/011)
-- ============================================================
-- Peticion de negocio: sustituir el texto heredado del legacy Java "Inclusion"/
-- "Exclusion" por "Bloqueo Terminal"/"Desbloqueo Terminal" en toda la aplicacion.
-- Esta tabla alimenta directamente el desplegable de alta de solicitud, el listado
-- y el detalle (los tres leen tipo_solicitud_ref.descripcion) - un UPDATE aqui
-- propaga el texto nuevo a los tres sitios sin tocar codigo. Se actualiza por CLAVE
-- ('I'/'E'), el identificador estable que usa el codigo, no por el texto que se
-- esta reemplazando. UPDATE simple sobre 2 filas como mucho, idempotente (reaplicar
-- con los valores ya puestos no cambia nada).
UPDATE tipo_solicitud_ref SET descripcion = 'Bloqueo Terminal' WHERE clave = 'I';
UPDATE tipo_solicitud_ref SET descripcion = 'Desbloqueo Terminal' WHERE clave = 'E';

-- ============================================================
-- FIN
-- Tras aplicar, verificar:
--   - SELECT * FROM auditoria_accion LIMIT 1; (tabla creada)
--   - DESCRIBE usuarios; (password VARCHAR(255), algoritmo_password, cod_usuario
--     VARCHAR(50), cod_indice_usuario con Extra=auto_increment)
--   - DESCRIBE solicitud; (enlace_externo presente)
--   - SELECT descripcion FROM tipo_solicitud_ref WHERE clave IN ('I','E'); debe
--     devolver 'Bloqueo Terminal' y 'Desbloqueo Terminal'.
--   - Confirmar que la app legacy Java sigue funcionando igual (login, bloqueo,
--     solicitudes) antes de dar la migracion por cerrada.
-- ============================================================
