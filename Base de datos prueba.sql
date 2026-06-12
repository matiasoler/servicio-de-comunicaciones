-- ===================================================================================
-- Script de Creación de Tablas para el Servicio de Comunicaciones
-- Motor: MySQL / MariaDB
-- Recomendación: Usar siempre el motor InnoDB para soportar claves foráneas.
-- ===================================================================================
-- Tabla central para identificar a las personas (clientes, etc.)
-- Es la base para asociar una venta a alguien.
CREATE TABLE IF NOT EXISTS personas (
    id_persona INT AUTO_INCREMENT PRIMARY KEY,
    nombre VARCHAR(100) NOT NULL,
    apellido VARCHAR(100) NOT NULL,
    dni VARCHAR(20) UNIQUE NOT NULL COMMENT 'UNIQUE para evitar personas duplicadas por DNI'
) ENGINE=InnoDB CHARACTER SET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
-- Tabla maestra de sucursales. Contiene la info de conexión.
-- Es el corazón de la comunicación.
CREATE TABLE IF NOT EXISTS sucursales_comunicacion (
    id_sucursal INT PRIMARY KEY,
    des_sucursal VARCHAR(255) NOT NULL,
    cadena_conexion VARCHAR(1024) NOT NULL COMMENT 'Aumentado por si las cadenas son largas',
    id_proceso INT,
    exceptuado BOOLEAN NOT NULL DEFAULT FALSE COMMENT 'TRUE si esta sucursal no participa en la comunicación'
) ENGINE=InnoDB CHARACTER SET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
-- Tabla principal de transacciones.
CREATE TABLE IF NOT EXISTS comprobantes_de_ventas (
    id_sucursal INT,
    id_comprobante BIGINT,
    total DECIMAL(15, 4) NOT NULL COMMENT 'Usar DECIMAL para moneda, nunca FLOAT o DOUBLE',
    persona_asociada INT NULL,
    fecha_comprobante TIMESTAMP DEFAULT CURRENT_TIMESTAMP, -- Es buena idea registrar cuándo se hizo.
    
    PRIMARY KEY (id_sucursal, id_comprobante),
    FOREIGN KEY (id_sucursal) REFERENCES sucursales_comunicacion(id_sucursal) ON DELETE CASCADE,
    FOREIGN KEY (persona_asociada) REFERENCES personas(id_persona) ON DELETE SET NULL
) ENGINE=InnoDB CHARACTER SET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
-- Tabla de parámetros. Ojo con la estructura.
-- Parece una tabla de configuración general.
CREATE TABLE IF NOT EXISTS parametros (
    id_empresa INT,
    des_empresa VARCHAR(255),
    id_sucursal INT,
    des_sucursal VARCHAR(255),
    dato_que_no_deberia_cambiar VARCHAR(255) COMMENT 'OJO: Un nombre de columna así es una MALA práctica. Debería ser específico, ej: version_schema, api_key_fija, etc.',
    
    PRIMARY KEY (id_empresa, id_sucursal) -- Asumo esta clave compuesta.
) ENGINE=InnoDB CHARACTER SET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
-- Tablas que se excluyen del proceso de sincronización.
CREATE TABLE IF NOT EXISTS tablas_exceptuadas_sucursal (
    id_tabla INT AUTO_INCREMENT PRIMARY KEY,
    tabla VARCHAR(128) NOT NULL,
    -- PREGUNTA: ¿La excepción es global o por sucursal?
    -- Si es por sucursal, deberías agregar un id_sucursal aquí.
    -- ej: id_sucursal INT, FOREIGN KEY (id_sucursal) REFERENCES sucursales_comunicacion(id_sucursal)
    UNIQUE (tabla) -- Asumo que el nombre de la tabla es único en esta lista.
) ENGINE=InnoDB CHARACTER SET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
-- Guardan la "posición" o el punto hasta donde se leyó un log de datos.
CREATE TABLE IF NOT EXISTS sucursales_log_recepcion (
    id_sucursal INT PRIMARY KEY,
    archivo_log VARCHAR(255) NOT NULL,
    posicion BIGINT NOT NULL DEFAULT 0 COMMENT 'BIGINT para archivos de log muy grandes',
    fecha_actualizacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    FOREIGN KEY (id_sucursal) REFERENCES sucursales_comunicacion(id_sucursal) ON DELETE CASCADE
) ENGINE=InnoDB CHARACTER SET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
-- Idem al anterior, pero para los logs de envío.
CREATE TABLE IF NOT EXISTS sucursales_log_envio (
    id_sucursal INT PRIMARY KEY,
    archivo_log VARCHAR(255) NOT NULL,
    posicion BIGINT NOT NULL DEFAULT 0,
    fecha_actualizacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
    FOREIGN KEY (id_sucursal) REFERENCES sucursales_comunicacion(id_sucursal) ON DELETE CASCADE
) ENGINE=InnoDB CHARACTER SET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
-- Registra métricas de tiempo para monitorear la performance.
CREATE TABLE IF NOT EXISTS sucursales_log_tiempo_comunicacion (
    id_log INT AUTO_INCREMENT PRIMARY KEY, -- Un log debería tener su propio ID.
    id_sucursal INT,
    tiempo_envio_ms INT COMMENT 'Duración en milisegundos',
    tiempo_recepcion_ms INT COMMENT 'Duración en milisegundos',
    fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (id_sucursal) REFERENCES sucursales_comunicacion(id_sucursal) ON DELETE CASCADE,
    INDEX idx_fecha_registro (fecha_registro) -- Indexar por fecha es clave en logs.
) ENGINE=InnoDB CHARACTER SET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
-- Log genérico. Le agregué una PK y un timestamp, que son fundamentales.
CREATE TABLE IF NOT EXISTS log_comunicaciones_externo (
    id_log BIGINT AUTO_INCREMENT PRIMARY KEY,
    id_sucursal INT NOT NULL,
    archivo_log VARCHAR(255) NOT NULL,
    pos BIGINT NOT NULL,
    mensaje TEXT, -- Un campo para describir el evento de log es útil.
    fecha_registro TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (id_sucursal) REFERENCES sucursales_comunicacion(id_sucursal) ON DELETE CASCADE,
    INDEX idx_fecha_sucursal (id_sucursal, fecha_registro)
) ENGINE=InnoDB CHARACTER SET=utf8mb4 COLLATE=utf8mb4_unicode_ci;
-- ===================================================================================
-- Fin del Script
-- ===================================================================================


-- 1. Añadimos la columna para marcar la sucursal central
ALTER TABLE sucursales_comunicacion ADD COLUMN es_central BOOLEAN NOT NULL DEFAULT FALSE;
-- 2. Creamos un índice en esa columna para que la búsqueda sea ultra rápida
CREATE INDEX idx_es_central ON sucursales_comunicacion(es_central);
/* 
   3. (IMPORTANTE) Ahora, tenés que marcar UNA y SÓLO UNA sucursal como central.
      Por ejemplo, si la central es la sucursal con id 1:
   UPDATE sucursales_comunicacion SET es_central = 1 WHERE id_sucursal = 1;
*/