-- =====================================================================
-- crear_bd.sql
-- Biblioteca Digital de Instrucciones de Trabajo - Impresos Múltiples
-- Fase 3: MySQL (XAMPP + phpMyAdmin) con documentos en LONGBLOB
--
-- Cómo usarlo: en phpMyAdmin → pestaña "Importar" → elegir este archivo → Continuar.
-- Si la base ya existe, borra y vuelve a crear las tablas (los datos se reinician).
-- =====================================================================

-- XAMPP trae max_allowed_packet = 1M: un documento más grande corta la conexión.
-- Esta línea lo sube a 64 MB hasta que se reinicie MySQL. Para que sea permanente,
-- poner max_allowed_packet=64M en la sección [mysqld] de xampp/mysql/bin/my.ini.
SET GLOBAL max_allowed_packet = 67108864;  -- 64 MB (documentos de hasta 50 MB)

CREATE DATABASE IF NOT EXISTS biblioteca_impresos
    CHARACTER SET utf8mb4
    COLLATE utf8mb4_unicode_ci;

USE biblioteca_impresos;

-- Se borran primero las tablas relacionadas (las que tienen llave foránea)
DROP TABLE IF EXISTS pasos;
DROP TABLE IF EXISTS archivos;
DROP TABLE IF EXISTS historial_versiones;
DROP TABLE IF EXISTS instrucciones;

-- ---------------------------------------------------------------------
-- Tabla principal: instrucciones de trabajo
-- ---------------------------------------------------------------------
CREATE TABLE instrucciones (
    id                  INT AUTO_INCREMENT PRIMARY KEY,
    codigo              VARCHAR(20)  NOT NULL UNIQUE,
    titulo              VARCHAR(100) NOT NULL,
    area                VARCHAR(50)  NOT NULL,
    enlace_video        VARCHAR(255) NULL,     -- video de la instrucción (solo el enlace)
    fecha_actualizacion DATE         NOT NULL,
    activa              TINYINT(1)   NOT NULL DEFAULT 1,  -- 1 = en uso, 0 = archivada
    fecha_archivado     DATETIME     NULL                -- cuándo se archivó
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ---------------------------------------------------------------------
-- Tabla relacionada: pasos de cada instrucción (el contenido que se edita
-- desde Registrar y Modificar). Se numeran 1, 2, 3... en el orden en que
-- el operario los realiza.
-- ON DELETE CASCADE: los pasos son parte de la propia instrucción (no son
-- historial). Una instrucción solo se puede eliminar si no tiene documentos;
-- en ese caso sus pasos se van con ella.
-- ---------------------------------------------------------------------
CREATE TABLE pasos (
    id             INT AUTO_INCREMENT PRIMARY KEY,
    instruccion_id INT          NOT NULL,
    numero         INT          NOT NULL,      -- orden del paso: 1, 2, 3...
    descripcion    VARCHAR(500) NOT NULL,      -- qué hacer
    punto_control  VARCHAR(255) NULL,          -- qué verificar o qué precaución tomar
    CONSTRAINT uq_paso_numero UNIQUE (instruccion_id, numero),
    CONSTRAINT fk_paso_instruccion
        FOREIGN KEY (instruccion_id) REFERENCES instrucciones (id)
        ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ---------------------------------------------------------------------
-- Tabla relacionada: historial de versiones de cada instrucción.
-- Una fila por cada versión del documento que se crea (subida o restauración).
-- instruccion_id es la LLAVE FORÁNEA: apunta a instrucciones.id
-- ON DELETE RESTRICT: NO se puede borrar una instrucción que tenga historial.
-- Así el historial nunca se pierde por eliminar un registro.
-- ---------------------------------------------------------------------
CREATE TABLE historial_versiones (
    id             INT AUTO_INCREMENT PRIMARY KEY,
    instruccion_id INT          NOT NULL,
    version        INT          NOT NULL,
    accion         VARCHAR(255) NOT NULL DEFAULT 'Versión registrada',
    fecha          DATETIME     NOT NULL,
    CONSTRAINT fk_historial_instruccion
        FOREIGN KEY (instruccion_id) REFERENCES instrucciones (id)
        ON DELETE RESTRICT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ---------------------------------------------------------------------
-- Tabla de documentos (Fase 3): un documento PDF o Word por instrucción,
-- con historial de versiones. Subir de nuevo el documento crea una fila
-- nueva con version = MAX(version) + 1; las anteriores no se tocan.
-- instruccion_id es la LLAVE FORÁNEA: apunta a instrucciones.id
-- ON DELETE RESTRICT: MySQL rechaza borrar una instrucción que tenga
-- documentos, así que nunca quedan archivos huérfanos ni se pierden versiones.
-- ---------------------------------------------------------------------
CREATE TABLE archivos (
    id              INT AUTO_INCREMENT PRIMARY KEY,
    instruccion_id  INT          NOT NULL,
    version         INT          NOT NULL,          -- 1, 2, 3... por instrucción
    nombre_original VARCHAR(255) NOT NULL,          -- nombre con el que se subió
    extension       VARCHAR(5)   NOT NULL,          -- pdf, doc o docx
    tamano_bytes    INT          NOT NULL,
    contenido       LONGBLOB     NOT NULL,          -- el archivo completo
    enlace_video    VARCHAR(255) NULL,              -- enlace del video al subir esta versión
    nota_cambio     VARCHAR(300) NULL,              -- qué cambió respecto a la versión anterior
    fecha_subida    DATETIME     NOT NULL,
    CONSTRAINT uq_archivo_version UNIQUE (instruccion_id, version),
    CONSTRAINT fk_archivo_instruccion
        FOREIGN KEY (instruccion_id) REFERENCES instrucciones (id)
        ON DELETE RESTRICT
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

-- ---------------------------------------------------------------------
-- Datos de ejemplo: 8 instrucciones (sin documentos: los documentos se
-- suben desde el navegador; el script no guarda archivos reales)
-- ---------------------------------------------------------------------
INSERT INTO instrucciones (id, codigo, titulo, area, enlace_video, fecha_actualizacion) VALUES
    (1, 'IM-OFF-001', 'Preparación de máquina offset', 'Impresión Offset', 'https://intranet.ejemplo.com/videos/offset-preparacion', '2026-09-02'),
    (2, 'IM-OFF-002', 'Cambio de planchas', 'Impresión Offset', 'https://intranet.ejemplo.com/videos/cambio-planchas', '2026-07-15'),
    (3, 'IM-TIN-001', 'Preparación de tinta Pantone', 'Preparación de Tintas', NULL, '2026-09-10'),
    (4, 'IM-TRO-001', 'Ajuste de troquel', 'Troquelado', 'https://intranet.ejemplo.com/videos/ajuste-troquel', '2026-05-20'),
    (5, 'IM-ACA-001', 'Laminado brillante', 'Acabados', NULL, '2026-08-01'),
    (6, 'IM-ENC-001', 'Encuadernado hot melt', 'Encuadernación', 'https://intranet.ejemplo.com/videos/hot-melt', '2026-04-11'),
    (7, 'IM-MAN-001', 'Calibración de cortadora', 'Mantenimiento', NULL, '2026-09-25'),
    (8, 'IM-CAL-001', 'Inspección de calidad de impresión', 'Control de Calidad', 'https://intranet.ejemplo.com/videos/inspeccion-calidad', '2026-08-18');

-- ---------------------------------------------------------------------
-- Datos de ejemplo: pasos de cada instrucción (contenido inventado para
-- pruebas académicas; la empresa los reemplaza por sus pasos reales)
-- ---------------------------------------------------------------------
INSERT INTO pasos (instruccion_id, numero, descripcion, punto_control) VALUES
    (1, 1, 'Revisar la orden de producción: tipo de papel, tintas y cantidad de pliegos.', 'La orden debe estar firmada por el supervisor.'),
    (1, 2, 'Limpiar rodillos y mantillas con el solvente indicado.', 'Usar guantes y lentes de protección.'),
    (1, 3, 'Montar las planchas en cada cuerpo de impresión y ajustar el registro.', NULL),
    (1, 4, 'Cargar el papel en el alimentador y regular la entrada de pliegos.', NULL),
    (1, 5, 'Imprimir 10 pliegos de prueba y compararlos con la prueba de color aprobada.', 'Densidad y registro dentro de tolerancia antes de iniciar el tiraje.'),
    (2, 1, 'Detener la máquina y bloquear el arranque.', 'Aplicar el bloqueo antes de tocar los cilindros.'),
    (2, 2, 'Retirar las planchas usadas y limpiar los cilindros portaplanchas.', NULL),
    (2, 3, 'Colocar las planchas nuevas siguiendo el orden de colores de la orden.', 'Verificar el código de cada plancha contra la orden.'),
    (2, 4, 'Ajustar el registro y hacer una prueba de 5 pliegos.', NULL),
    (3, 1, 'Consultar la fórmula del color Pantone en la guía de mezclas.', NULL),
    (3, 2, 'Pesar cada tinta base en la balanza según la fórmula.', 'Tolerancia de ±1 g por componente.'),
    (3, 3, 'Mezclar durante 3 minutos hasta obtener un color uniforme.', NULL),
    (3, 4, 'Hacer una prueba de color y compararla con la guía bajo luz estándar.', 'Registrar el resultado en la hoja de control.'),
    (4, 1, 'Colocar la forma de troquel en el portatroquel.', NULL),
    (4, 2, 'Ajustar la presión en tres puntos de la forma.', 'No superar la presión máxima indicada para el material.'),
    (4, 3, 'Verificar el corte y el hendido en 5 pliegos de prueba.', 'Sin rebabas ni cortes incompletos.'),
    (5, 1, 'Precalentar la laminadora a la temperatura indicada para el film.', NULL),
    (5, 2, 'Alinear el pliego impreso con la guía de entrada.', NULL),
    (5, 3, 'Laminar un pliego de prueba y revisar que no queden burbujas.', 'Sin burbujas ni arrugas en el acabado.'),
    (6, 1, 'Cargar el pegamento hot melt y esperar a que alcance la temperatura de trabajo.', 'Usar guantes térmicos.'),
    (6, 2, 'Ajustar el ancho del lomo según el grosor del libro.', NULL),
    (6, 3, 'Encuadernar 3 ejemplares de prueba y revisar la adherencia de las páginas.', NULL),
    (7, 1, 'Apagar la cortadora y bloquear el arranque.', 'Nunca calibrar con la máquina encendida.'),
    (7, 2, 'Medir la escuadra y el paralelismo de la cuchilla.', NULL),
    (7, 3, 'Hacer un corte de prueba y medir con el calibrador.', 'Tolerancia de ±0.5 mm.'),
    (8, 1, 'Tomar 5 muestras al azar del tiraje.', NULL),
    (8, 2, 'Revisar registro, color, texto y limpieza de cada muestra.', 'Comparar contra la prueba de color aprobada.'),
    (8, 3, 'Registrar el resultado y liberar o retener el lote.', 'Un lote retenido se identifica con etiqueta roja.');
