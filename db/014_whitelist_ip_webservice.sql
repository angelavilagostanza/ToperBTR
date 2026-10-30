-- 014 - Lista blanca de IPs para el nuevo webservice de intercambio (BTR-WS)
-- ============================================================
-- BTR-WS es un servicio nuevo, fuera de este repositorio (ver
-- docs/17-arquitectura-webservice-intercambio.md y
-- docs/18-plan-accion-webservice-intercambio.md), que sustituira a
-- app/externos/intercambio/descargar.asp para que los sistemas de otros operadores
-- descarguen el fichero diario de intercambio EIR. Esta tabla vive en la misma BD
-- compartida pre/PRO (eir_test en 172.30.170.27) porque es el unico almacen de datos
-- comun a ToperBTR y a BTR-WS - permite que, en una iteracion futura, se gestione
-- desde una pantalla de administracion en ToperBTR sin duplicar infraestructura.
--
-- Sin FK a `usuarios` a proposito: BTR-WS no debe depender del esquema interno de
-- ToperBTR evolucionando con el tiempo. Las columnas *_usuario son solo informativas,
-- rellenadas por la futura pantalla de administracion (Paso 7 del plan de accion).
--
-- Idempotente: CREATE TABLE IF NOT EXISTS, igual que el resto de scripts de este
-- directorio.
-- ============================================================

CREATE TABLE IF NOT EXISTS ws_ip_permitida (
  cod_ip_permitida  INT NOT NULL AUTO_INCREMENT,
  ip_o_rango        VARCHAR(43) NOT NULL,   -- IPv4/IPv6 exacta; deja margen para CIDR futuro
  descripcion       VARCHAR(200) NULL,      -- p.ej. "Operador X - salida fija oficina"
  activo            TINYINT(1) NOT NULL DEFAULT 1,
  fecha_alta        DATETIME NOT NULL,
  cod_usuario_alta  INT NULL,               -- referencia informativa a usuarios.cod_indice_usuario
  fecha_baja        DATETIME NULL,
  cod_usuario_baja  INT NULL,
  PRIMARY KEY (cod_ip_permitida),
  KEY idx_ip_activo (ip_o_rango, activo)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

-- ============================================================
-- FIN
-- Tras aplicar, verificar:
--   - DESCRIBE ws_ip_permitida;
--   - Insertar manualmente las primeras IPs conocidas de operadores para las
--     pruebas del Paso 4 del plan de accion, por ejemplo:
--     INSERT INTO ws_ip_permitida (ip_o_rango, descripcion, activo, fecha_alta)
--     VALUES ('203.0.113.10', 'Operador X - prueba inicial', 1, NOW());
-- ============================================================
