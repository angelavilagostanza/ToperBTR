-- Entregable 1 (Base + Login + Autenticacion + Autorizacion) - ToperBTR
-- Se aplica en vivo sobre la BD copia de produccion `eir_test` (no hay migracion a una BD nueva,
-- ver docs/08-cambios-bbdd.md). Cambios aditivos, pensados para convivir con la app legacy Java durante el periodo de convivencia acordado.
--
-- REVISAR ANTES DE EJECUTAR: los pasos marcados con -- VERIFICAR requieren comprobar
-- el estado real de la BD de produccion antes de aplicarlos (puede que ya existan).

-- 1) Ampliar columnas de password para el nuevo formato PBKDF2SHA256 (autodescriptivo,
--    con salt e iteraciones incluidos). El legacy sigue escribiendo valores cortos
--    (SHA1 + Base64) sin ningun problema en una columna mas ancha.
ALTER TABLE usuarios MODIFY COLUMN password VARCHAR(255) NOT NULL;
ALTER TABLE historico_password MODIFY COLUMN password VARCHAR(255) NOT NULL;

-- 2) Columna discriminadora de algoritmo, para que legacy (SHA1B64) y ToperBTR
--    (PBKDF2SHA256) convivan sin ambiguedad.
ALTER TABLE usuarios ADD COLUMN algoritmo_password VARCHAR(20) NOT NULL DEFAULT 'SHA1B64';
UPDATE usuarios SET algoritmo_password = 'SHA1B64' WHERE algoritmo_password = '';

-- 3) VERIFICAR: el codigo legacy ya referencia activamente UsuarioVO.BLOQUEADO=2,  es posible que esta fila ya exista. Ejecutar antes: SELECT * FROM estado_usuario_ref;
INSERT INTO estado_usuario_ref (estado, descripcion)
SELECT 2, 'Bloqueado'
WHERE NOT EXISTS (SELECT 1 FROM estado_usuario_ref WHERE estado = 2);

-- 4) VERIFICAR primero los valores actuales: SELECT * FROM parametros_ref WHERE clave IN ('NUM_FALLOS_LOGIN_USUARIO','NUM_OLD_PASSWORD','DURACION_PASSWORD','COD_OPER_HOME');
--    Esas 4 claves ya las usa el legacy, no se tocan aqui. Solo se anade la nueva:
INSERT INTO parametros_ref (clave, valor)
SELECT 'NUM_MAX_INTENTOS_FALLIDOS', '5'
WHERE NOT EXISTS (SELECT 1 FROM parametros_ref WHERE clave = 'NUM_MAX_INTENTOS_FALLIDOS');

-- 5) AUTO_INCREMENT en usuarios (cod_indice_usuario ya es PRIMARY KEY). MySQL fija
--    automaticamente el contador al MAX(cod_indice_usuario) existente al hacer el ALTER.
--    VERIFICAR antes si hay colisiones/huecos por la condicion de carrera del MAX+1 legacy:
--    SELECT cod_indice_usuario, COUNT(*) FROM usuarios GROUP BY cod_indice_usuario HAVING COUNT(*) > 1;
ALTER TABLE usuarios MODIFY COLUMN cod_indice_usuario INT NOT NULL AUTO_INCREMENT;

-- 6) Tabla nueva de auditoria (aditiva, sin relacion con tablas del legacy).
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
