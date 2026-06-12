import aiomysql
import logging
from urllib.parse import urlparse

logger = logging.getLogger(__name__)

class DatabaseManager:
    """
    Gestiona las conexiones a las bases de datos local y central.
    Es responsable de cargar la configuración dinámica y de persistir
    el estado de la replicación para ambos flujos (subida y bajada).
    """
    def __init__(self, static_config):
        self.static_config = static_config
        self.local_pool = None
        self.central_pool = None
        self.dynamic_config = {}
        self.upload_resume_position = {}
        self.download_resume_position = {}

    def _parse_dsn(self, dsn: str) -> dict:
        """Parsea una cadena de conexión DSN a un diccionario para aiomysql."""
        parsed = urlparse(dsn)
        if not all([parsed.scheme, parsed.hostname, parsed.username, parsed.path]):
            raise ValueError(f"La cadena de conexión (DSN) '{dsn}' es inválida o incompleta.")
        
        return {
            'user': parsed.username,
            'password': parsed.password,
            'host': parsed.hostname,
            'port': parsed.port or 3306,
            'db': parsed.path.lstrip('/')
        }

    async def initialize(self):
        """
        Establece la conexión con ambas bases de datos (local y central)
        y carga la configuración inicial de replicación para ambos flujos.
        """
        logger.info("Inicializando DatabaseManager...")
        try:
            db_settings = self.static_config['database_local']
            self.local_pool = await aiomysql.create_pool(
                host=db_settings['host'], port=int(db_settings['port']),
                user=db_settings['user'], password=db_settings['password'],
                db=db_settings['db'], autocommit=False
            )
            logger.info("Pool de conexiones a la base de datos LOCAL creado.")

            await self._load_dynamic_config_and_central_pool()
            await self._load_all_resume_positions()

        except Exception as e:
            logger.error(f"No se pudo inicializar DatabaseManager: {e}", exc_info=True)
            await self.close()
            raise

    async def _load_dynamic_config_and_central_pool(self):
        """Lee la tabla de sucursales para encontrar la BD central y conectarse a ella."""
        logger.info("Cargando configuración dinámica desde la BD local...")
        async with self.local_pool.acquire() as conn:
            async with conn.cursor(aiomysql.DictCursor) as cursor:
                await cursor.execute("SELECT cadena_conexion FROM sucursales_comunicacion WHERE es_central = TRUE LIMIT 1")
                result = await cursor.fetchone()
                if not result or not result.get('cadena_conexion'):
                    raise Exception("No se encontró la configuración de la sucursal central ('es_central = TRUE') en la tabla 'sucursales_comunicacion'.")
                
                central_db_dsn = result['cadena_conexion']
                central_db_settings = self._parse_dsn(central_db_dsn)
                self.dynamic_config['database_central'] = central_db_settings
                
                self.central_pool = await aiomysql.create_pool(**central_db_settings, autocommit=False)
                logger.info("Pool de conexiones a la base de datos CENTRAL creado.")

    async def _load_all_resume_positions(self):
        """Carga las posiciones de replicación para los flujos de subida y bajada."""
        branch_id = self.static_config.getint('general', 'branch_id')
        async with self.local_pool.acquire() as conn:
            async with conn.cursor(aiomysql.DictCursor) as cursor:
                # Cargar posición de SUBIDA
                await cursor.execute("SELECT archivo_log, posicion FROM sucursales_log_envio WHERE id_sucursal = %s", (branch_id,))
                result_upload = await cursor.fetchone()
                if result_upload and result_upload.get('archivo_log'):
                    self.upload_resume_position = {'log_file': result_upload['archivo_log'], 'log_pos': result_upload['posicion']}
                else:
                    logger.warning(f"No se encontró posición de SUBIDA para sucursal {branch_id}. Obteniendo estado del master LOCAL.")
                    await cursor.execute("SHOW MASTER STATUS")
                    status = await cursor.fetchone()
                    if not status: raise Exception("No se pudo obtener el 'MASTER STATUS' de la BD local.")
                    self.upload_resume_position = {'log_file': status['File'], 'log_pos': status['Position']}
                    await self.save_upload_position(status['File'], status['Position'])
                logger.info(f"Posición de SUBIDA (local -> central) establecida: {self.upload_resume_position}")

                # Cargar posición de BAJADA
                await cursor.execute("SELECT archivo_log, posicion FROM sucursales_log_recepcion WHERE id_sucursal = %s", (branch_id,))
                result_download = await cursor.fetchone()
                if result_download and result_download.get('archivo_log'):
                    self.download_resume_position = {'log_file': result_download['archivo_log'], 'log_pos': result_download['posicion']}
                else:
                    logger.warning(f"No se encontró posición de BAJADA para sucursal {branch_id}. Obteniendo estado del master CENTRAL.")
                    async with self.central_pool.acquire() as central_conn:
                        async with central_conn.cursor(aiomysql.DictCursor) as central_cursor:
                            await central_cursor.execute("SHOW MASTER STATUS")
                            status = await central_cursor.fetchone()
                            if not status: raise Exception("No se pudo obtener el 'MASTER STATUS' de la BD central.")
                            self.download_resume_position = {'log_file': status['File'], 'log_pos': status['Position']}
                            await self.save_download_position(status['File'], status['Position'])
                logger.info(f"Posición de BAJADA (central -> local) establecida: {self.download_resume_position}")

    def get_local_pool(self):
        return self.local_pool

    def get_central_pool(self):
        return self.central_pool

    def get_upload_resume_position(self):
        return self.upload_resume_position

    def get_download_resume_position(self):
        return self.download_resume_position

    async def save_upload_position(self, log_file: str, log_pos: int):
        """Guarda la última posición del binlog de SUBIDA procesada."""
        sql = """
            INSERT INTO sucursales_log_envio (id_sucursal, archivo_log, posicion, fecha_actualizacion)
            VALUES (%s, %s, %s, NOW())
            ON DUPLICATE KEY UPDATE archivo_log = %s, posicion = %s, fecha_actualizacion = NOW()
        """
        await self._save_position(sql, log_file, log_pos)
        self.upload_resume_position = {'log_file': log_file, 'log_pos': log_pos}

    async def save_download_position(self, log_file: str, log_pos: int):
        """Guarda la última posición del binlog de BAJADA procesada."""
        sql = """
            INSERT INTO sucursales_log_recepcion (id_sucursal, archivo_log, posicion, fecha_actualizacion)
            VALUES (%s, %s, %s, NOW())
            ON DUPLICATE KEY UPDATE archivo_log = %s, posicion = %s, fecha_actualizacion = NOW()
        """
        await self._save_position(sql, log_file, log_pos)
        self.download_resume_position = {'log_file': log_file, 'log_pos': log_pos}

    async def _save_position(self, sql: str, log_file: str, log_pos: int):
        """Método genérico para guardar una posición en la BD local."""
        branch_id = self.static_config.getint('general', 'branch_id')
        async with self.local_pool.acquire() as conn:
            async with conn.cursor() as cursor:
                try:
                    await conn.begin()
                    await cursor.execute(sql, (branch_id, log_file, log_pos, log_file, log_pos))
                    await conn.commit()
                except Exception as e:
                    logger.error(f"Error al guardar posición para sucursal {branch_id}: {e}", exc_info=True)
                    await conn.rollback()
                    raise

    async def close(self):
        """Cierra todos los pools de conexiones a las bases de datos."""
        if self.local_pool:
            self.local_pool.close()
            await self.local_pool.wait_closed()
            logger.info("Pool de conexiones a BD local cerrado.")
        if self.central_pool:
            self.central_pool.close()
            await self.central_pool.wait_closed()
            logger.info("Pool de conexiones a BD central cerrado.")

