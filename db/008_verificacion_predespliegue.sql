-- Verificaciones previas al despliegue de ToperBTR en eir_test
-- ============================================================
-- Ejecutar ANTES de aplicar los scripts 001-007.
-- Son consultas SELECT-only: no modifican nada.
-- Cada seccion indica que se espera y que indica un problema.
--
-- Si alguna consulta devuelve resultados inesperados, revisar
-- docs/08-cambios-bbdd.md antes de continuar.

-- ============================================================
-- 1) TABLAS REQUERIDAS POR LOS SCRIPTS 001-007
--    Deben existir en la BD de produccion antes de aplicar los scripts.
-- ============================================================

SELECT 'TABLA_USUARIOS' AS verificacion,
       CASE WHEN COUNT(*) > 0 THEN 'OK' ELSE 'ERROR: tabla no encontrada' END AS resultado
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'usuarios'

UNION ALL

SELECT 'TABLA_HISTORICO_PASSWORD',
       CASE WHEN COUNT(*) > 0 THEN 'OK' ELSE 'ERROR: tabla no encontrada' END
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'historico_password'

UNION ALL

SELECT 'TABLA_ESTADO_USUARIO_REF',
       CASE WHEN COUNT(*) > 0 THEN 'OK' ELSE 'ERROR: tabla no encontrada' END
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'estado_usuario_ref'

UNION ALL

SELECT 'TABLA_PARAMETROS_REF',
       CASE WHEN COUNT(*) > 0 THEN 'OK' ELSE 'ERROR: tabla no encontrada' END
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'parametros_ref'

UNION ALL

SELECT 'TABLA_PERFIL_REF',
       CASE WHEN COUNT(*) > 0 THEN 'OK' ELSE 'ERROR: tabla no encontrada' END
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'perfil_ref'

UNION ALL

SELECT 'TABLA_ACCION_REF',
       CASE WHEN COUNT(*) > 0 THEN 'OK' ELSE 'ERROR: tabla no encontrada' END
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'accion_ref'

UNION ALL

SELECT 'TABLA_ACCION_PERMITIDA_REF',
       CASE WHEN COUNT(*) > 0 THEN 'OK' ELSE 'ERROR: tabla no encontrada' END
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'accion_permitida_ref'

UNION ALL

SELECT 'TABLA_LISTA_NEGRA',
       CASE WHEN COUNT(*) > 0 THEN 'OK' ELSE 'ERROR: tabla no encontrada' END
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'lista_negra'

UNION ALL

SELECT 'TABLA_LISTA_BLANCA',
       CASE WHEN COUNT(*) > 0 THEN 'OK' ELSE 'ERROR: tabla no encontrada' END
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'lista_blanca'

UNION ALL

SELECT 'TABLA_HISTORICO_BLOQUEO_DIRECTO',
       CASE WHEN COUNT(*) > 0 THEN 'OK' ELSE 'ERROR: tabla no encontrada' END
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'historico_bloqueo_directo'

UNION ALL

SELECT 'TABLA_FICHEROS',
       CASE WHEN COUNT(*) > 0 THEN 'OK' ELSE 'ERROR: tabla no encontrada' END
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'ficheros'

UNION ALL

SELECT 'TABLA_SOLICITUDES',
       CASE WHEN COUNT(*) > 0 THEN 'OK' ELSE 'ERROR: tabla no encontrada' END
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'solicitudes'

UNION ALL

SELECT 'TABLA_INCIDENCIA_TO_OBJETO',
       CASE WHEN COUNT(*) > 0 THEN 'OK' ELSE 'ERROR: tabla no encontrada' END
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'incidencia_to_objeto'

UNION ALL

SELECT 'TABLA_TIPO_BLOQUEO_REF',
       CASE WHEN COUNT(*) > 0 THEN 'OK' ELSE 'ERROR: tabla no encontrada' END
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'tipo_bloqueo_ref'

UNION ALL

SELECT 'TABLA_RAZON_LISTA_NEGRA_REF',
       CASE WHEN COUNT(*) > 0 THEN 'OK' ELSE 'ERROR: tabla no encontrada' END
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'razon_lista_negra_ref'
;

-- ============================================================
-- 2) COLUMNAS CRITICAS EN USUARIOS (usadas en autenticacion y
--    modificadas por script 001)
-- ============================================================

SELECT 'COL_USUARIOS_PASSWORD' AS verificacion,
       CASE WHEN COUNT(*) > 0 THEN 'OK' ELSE 'ERROR: columna no encontrada' END AS resultado
FROM information_schema.COLUMNS
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'usuarios' AND COLUMN_NAME = 'password'

UNION ALL

SELECT 'COL_USUARIOS_ALGORITMO_PASSWORD',
       CASE WHEN COUNT(*) > 0 THEN 'EXISTE (script 001 ya aplicado)' ELSE 'NO_EXISTE (pendiente, ok)' END
FROM information_schema.COLUMNS
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'usuarios' AND COLUMN_NAME = 'algoritmo_password'

UNION ALL

SELECT 'COL_HISTORICO_PASSWORD_ALGORITMO_PASSWORD',
       CASE WHEN COUNT(*) > 0 THEN 'EXISTE (script 002 ya aplicado)' ELSE 'NO_EXISTE (pendiente, ok)' END
FROM information_schema.COLUMNS
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'historico_password' AND COLUMN_NAME = 'algoritmo_password'
;

-- ============================================================
-- 3) FILAS DE CATALOGO REQUERIDAS POR TOPERBTR
--    Deben existir antes de que los modulos puedan resolver codigos por nombre.
-- ============================================================

-- Perfiles minimos esperados: Administrador, Seguridad, Tramitacion, Tercero
SELECT 'PERFILES_MINIMOS' AS verificacion,
       GROUP_CONCAT(NOMBRE ORDER BY NOMBRE) AS encontrados,
       CASE
           WHEN SUM(NOMBRE IN ('Administrador','Seguridad','Tramitación','Tercero')) >= 4
           THEN 'OK'
           ELSE 'REVISAR: pueden faltar perfiles'
       END AS resultado
FROM perfil_ref
;

-- Estado 'Bloqueado' (cod 2) — script 001 lo inserta si no existe
SELECT 'ESTADO_USUARIO_BLOQUEADO' AS verificacion,
       CASE WHEN COUNT(*) > 0 THEN 'OK' ELSE 'NO_EXISTE (pendiente, lo crea script 001)' END AS resultado
FROM estado_usuario_ref
WHERE estado = 2 AND descripcion = 'Bloqueado'
;

-- Parametros del legacy que deben preexistir (los 4 originales)
SELECT 'PARAMETROS_LEGACY' AS verificacion,
       GROUP_CONCAT(clave ORDER BY clave) AS encontrados,
       CASE
           WHEN SUM(clave IN ('COD_OPER_HOME','DIAS_MARGEN','DURACION_PASSWORD',
                              'NUM_FALLOS_LOGIN_USUARIO','NUM_OLD_PASSWORD')) >= 4
           THEN 'OK (al menos 4 de 5 legacy)'
           ELSE 'REVISAR: pueden faltar parametros del legacy'
       END AS resultado
FROM parametros_ref
;

-- Parametro nuevo de ToperBTR — script 001 lo inserta si no existe
SELECT 'PARAMETRO_NUM_MAX_INTENTOS_FALLIDOS' AS verificacion,
       CASE WHEN COUNT(*) > 0 THEN 'EXISTE (script 001 ya aplicado)' ELSE 'NO_EXISTE (pendiente, ok)' END AS resultado
FROM parametros_ref
WHERE clave = 'NUM_MAX_INTENTOS_FALLIDOS'
;

-- Tipos de bloqueo: deben existir al menos BLOQUEO DIRECTO y BLOQUEO GLOBAL
SELECT 'TIPOS_BLOQUEO_MINIMOS' AS verificacion,
       GROUP_CONCAT(descripcion ORDER BY descripcion) AS encontrados,
       CASE
           WHEN COUNT(*) >= 2 THEN 'OK'
           ELSE 'REVISAR: necesarios al menos BLOQUEO DIRECTO y BLOQUEO GLOBAL'
       END AS resultado
FROM tipo_bloqueo_ref
;

-- ============================================================
-- 4) TABLA AUDITORIA_ACCION (crea el script 001 si no existe)
-- ============================================================

SELECT 'TABLA_AUDITORIA_ACCION' AS verificacion,
       CASE WHEN COUNT(*) > 0 THEN 'EXISTE (script 001 ya aplicado)' ELSE 'NO_EXISTE (pendiente, ok)' END AS resultado
FROM information_schema.TABLES
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'auditoria_accion'
;

-- ============================================================
-- 5) AUTO_INCREMENT en tablas que lo requieren
--    (verificar que los scripts no fallan por ausencia de AI)
-- ============================================================

SELECT 'AI_USUARIOS' AS verificacion,
       CASE WHEN EXTRA LIKE '%auto_increment%' THEN 'OK' ELSE 'PENDIENTE (script 001 lo activa)' END AS resultado
FROM information_schema.COLUMNS
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'usuarios' AND COLUMN_NAME = 'cod_indice_usuario'

UNION ALL

SELECT 'AI_LISTA_NEGRA',
       CASE WHEN EXTRA LIKE '%auto_increment%' THEN 'OK' ELSE 'PENDIENTE (script 006 lo activa)' END
FROM information_schema.COLUMNS
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'lista_negra' AND COLUMN_NAME = 'cod_lista_negra'

UNION ALL

SELECT 'AI_HISTORICO_BLOQUEO_DIRECTO',
       CASE WHEN EXTRA LIKE '%auto_increment%' THEN 'OK' ELSE 'PENDIENTE (script 003 lo activa)' END
FROM information_schema.COLUMNS
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'historico_bloqueo_directo' AND COLUMN_NAME = 'cod_hist_bloq_dir'

UNION ALL

SELECT 'AI_CLIENTE',
       CASE WHEN EXTRA LIKE '%auto_increment%' THEN 'OK' ELSE 'PENDIENTE (script 003 lo activa)' END
FROM information_schema.COLUMNS
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'cliente' AND COLUMN_NAME = 'cod_cliente'

UNION ALL

SELECT 'AI_SOLICITUD',
       CASE WHEN EXTRA LIKE '%auto_increment%' THEN 'OK' ELSE 'PENDIENTE (script 003 lo activa)' END
FROM information_schema.COLUMNS
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'solicitud' AND COLUMN_NAME = 'cod_indice_sol'

UNION ALL

SELECT 'AI_COMENTARIO_INCIDENCIA',
       CASE WHEN EXTRA LIKE '%auto_increment%' THEN 'OK' ELSE 'PENDIENTE (script 004 lo activa)' END
FROM information_schema.COLUMNS
WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'comentario_incidencia' AND COLUMN_NAME = 'cod_indice_comentario'
;

-- ============================================================
-- 6) POSIBLES DUPLICADOS EN USUARIOS (race condition del legacy)
--    Si aparecen filas con COUNT > 1, resolver ANTES de aplicar
--    el script 001 (que pone AUTO_INCREMENT en cod_indice_usuario).
-- ============================================================

SELECT 'DUPLICADOS_COD_INDICE_USUARIO' AS verificacion,
       CASE WHEN COUNT(*) = 0 THEN 'OK (sin duplicados)' ELSE 'ERROR: hay duplicados, resolver antes del despliegue' END AS resultado
FROM (
    SELECT cod_indice_usuario
    FROM usuarios
    GROUP BY cod_indice_usuario
    HAVING COUNT(*) > 1
) t
;

-- ============================================================
-- FIN DE LA VERIFICACION
-- Todos los resultados con 'ERROR' deben resolverse antes de aplicar
-- los scripts de migracion 001-007.
-- Los resultados 'PENDIENTE' y 'NO_EXISTE' son normales antes de aplicar
-- el script correspondiente.
-- ============================================================
