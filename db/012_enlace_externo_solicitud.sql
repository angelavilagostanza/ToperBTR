-- 012 - Anadir ENLACE_EXTERNO a SOLICITUD - ToperBTR
-- ============================================================
-- Contexto (2026-10-29): peticion de negocio - guardar un enlace a otra aplicacion
-- (Jira, CRM, etc.) asociado a cada solicitud. Se rellena al dar de alta la solicitud
-- (app/modules/solicitudes/alta.asp/alta_do.asp) y se muestra de solo lectura en el
-- detalle (app/modules/solicitudes/detalle.asp) - no hay pantalla de edicion de
-- solicitudes en ToperBTR hoy, asi que no se puede modificar despues de creada.
--
-- VARCHAR(500): mismo tamaño que ya usa el campo Comentario en otros modulos del
-- proyecto (bloqueos/bloquear.asp, incidencias/detalle.asp, admin/parametros.asp) -
-- de sobra para una URL normal (un ejemplo real dado: https://masorange.atlassian.net/
-- browse/B2BSF-2918, unos 55 caracteres) con margen para enlaces mas largos o con
-- parametros.
--
-- NULL por defecto (no NOT NULL): es un dato opcional, las solicitudes ya existentes no
-- tienen ningun valor que migrar aqui.
--
-- ADD COLUMN protegido con comprobacion previa de information_schema (mismo patron que
-- db/009_reaplicacion_idempotente.sql) para poder reaplicar este script sin error si ya
-- se hubiera ejecutado antes.

USE eir_test;

SET @existe_columna = (
    SELECT COUNT(*) FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'solicitud' AND COLUMN_NAME = 'enlace_externo'
);

SET @sql_alter = IF(@existe_columna = 0,
    'ALTER TABLE solicitud ADD COLUMN enlace_externo VARCHAR(500) NULL',
    'SELECT 1'
);

PREPARE stmt FROM @sql_alter;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;
