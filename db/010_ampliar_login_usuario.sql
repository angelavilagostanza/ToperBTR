-- 010 - Ampliar el login de usuario a 50 caracteres - ToperBTR
-- ============================================================
-- Contexto (2026-10-06): decision de negocio confirmada - los usuarios nuevos se dan
-- de alta de forma nominal y el login pasa a ser la cuenta de correo electronico, que
-- no cabe en el limite legacy de 15 caracteres. Se amplia tanto en el codigo (ver
-- app/modules/login/login.asp, app/modules/usuarios/alta.asp, alta_do.asp,
-- listar.asp) como aqui, en la columna real.
--
-- ALTER ... MODIFY COLUMN es idempotente por naturaleza en MySQL (reaplicar la misma
-- definicion no da error), igual que el resto de ALTER de este proyecto - seguro de
-- re-ejecutar.
--
-- IMPORTANTE: antes de aplicar, confirmar el tamaño real actual con:
--   SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH
--   FROM information_schema.COLUMNS
--   WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'usuarios' AND COLUMN_NAME = 'cod_usuario';
-- No hay ningun CREATE/ALTER TABLE previo de cod_usuario en los scripts 001-009 de este
-- repositorio - es columna legacy preexistente, su anchura actual no consta en control
-- de versiones.

USE eir_test;

ALTER TABLE usuarios MODIFY COLUMN cod_usuario VARCHAR(50) NOT NULL;
