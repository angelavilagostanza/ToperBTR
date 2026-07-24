-- Entregable 2 (Usuarios) - ToperBTR
-- Se aplica en vivo sobre la BD copia de produccion `eir_test` (ver docs/08-cambios-bbdd.md),
-- pensado para convivir con la app legacy Java. Cambio aditivo.

-- El historico de contrasenas puede contener, a partir de ahora, tanto valores legacy
-- (SHA1B64) como nuevos (PBKDF2SHA256) segun cuando se generase cada fila. Sin esta
-- columna no habria forma de saber, fila a fila, con que algoritmo comparar al comprobar
-- si una contrasena nueva ya se uso antes (ver PasswordReutilizada en dal/usuarios_dal.asp).
ALTER TABLE historico_password ADD COLUMN algoritmo_password VARCHAR(20) NOT NULL DEFAULT 'SHA1B64';
UPDATE historico_password SET algoritmo_password = 'SHA1B64' WHERE algoritmo_password = '';
