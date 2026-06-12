# database/connection.py
import aiomysql
import configparser

def get_local_db_config(config_path: str) -> dict:
    """Lee el config.ini y retorna el diccionario de configuracion local."""
    config = configparser.ConfigParser()
    config.read(config_path)
    db_config = {k: config.get('database_local', k) for k in config.options('database_local')}
    db_config['port'] = int(db_config['port'])
    db_config['autocommit'] = True
    return db_config

def get_replication_config(config_path: str) -> dict:
    config = configparser.ConfigParser()
    config.read(config_path)
    try:
        max_blocks = config.getint('replication', 'max_blocks_per_batch')
    except (configparser.NoSectionError, configparser.NoOptionError):
        max_blocks = 10
    return {'max_blocks_per_batch': max_blocks}

async def create_db_pool(config: dict) -> aiomysql.Pool:
    """Crea y retorna un pool de conexiones."""
    return await aiomysql.create_pool(**config)

async def discover_identity(pool: aiomysql.Pool) -> int:
    """Obtiene el id_sucursal de la tabla parametros."""
    async with pool.acquire() as conn:
        async with conn.cursor(aiomysql.DictCursor) as cursor:
            await cursor.execute("SELECT id_sucursal FROM parametros LIMIT 1")
            result = await cursor.fetchone()
            if not result:
                raise Exception("Error crítico: No se encontró la fila de parámetros en la DB local.")
            return result['id_sucursal']

async def discover_central(pool: aiomysql.Pool) -> tuple[int, dict]:
    """Descubre la sucursal central y retorna su ID y configuracion de conexion."""
    async with pool.acquire() as conn:
        async with conn.cursor(aiomysql.DictCursor) as cursor:
            await cursor.execute("SELECT id_sucursal, cadena_conexion FROM sucursales_comunicacion WHERE es_central = 1")
            centrales = await cursor.fetchall()
            
            if len(centrales) == 0:
                raise Exception("Error crítico: No se encontró ninguna sucursal marcada como central (es_central=1).")
            elif len(centrales) > 1:
                raise Exception(f"Error crítico: Se encontraron {len(centrales)} sucursales centrales. Se aborta la ejecución.")
            
            central_info = centrales[0]
            
            dsn = central_info['cadena_conexion']
            if not dsn:
                raise Exception(f"Error crítico: La sucursal central (id={central_info['id_sucursal']}) no tiene una 'cadena_conexion' configurada.")
            
            parts = dsn.replace("mysql://", "").split("@")
            user_pass, host_db = parts[0], parts[1]
            user, password = user_pass.split(":")
            host_port, db = host_db.split("/")
            host, port = (host_port.split(":") + [3306])[:2]
            central_config = {'host': host, 'port': int(port), 'user': user, 'password': password, 'db': db, 'autocommit': True}
            
            return central_info['id_sucursal'], central_config

async def get_excepted_tables(pool: aiomysql.Pool, target_id: int) -> set:
    """Obtiene la lista (set) de tablas exceptuadas para una sucursal destino."""
    async with pool.acquire() as conn:
        async with conn.cursor(aiomysql.DictCursor) as cursor:
            await cursor.execute("SELECT tabla FROM tablas_exceptuadas_sucursal WHERE id_sucursal = %s", (target_id,))
            return {row['tabla'] for row in await cursor.fetchall()}

async def verify_binlog_config(pool: aiomysql.Pool, db_name: str, host_ip: str) -> None:
    """Verifica que el binlog esté activo y en formato ROW."""
    async with pool.acquire() as conn:
        async with conn.cursor(aiomysql.DictCursor) as cursor:
            await cursor.execute("SHOW VARIABLES LIKE 'log_bin'")
            log_bin = await cursor.fetchone()
            if not log_bin or log_bin['Value'].upper() != 'ON':
                raise Exception(f"Error crítico: El binlog (log_bin) NO está activado en {db_name} (IP: {host_ip}).")
            
            await cursor.execute("SHOW VARIABLES LIKE 'binlog_format'")
            binlog_format = await cursor.fetchone()
            if not binlog_format or binlog_format['Value'].upper() != 'ROW':
                raise Exception(f"Error crítico: El binlog_format en {db_name} (IP: {host_ip}) debe ser ROW. Actual: {binlog_format['Value'] if binlog_format else 'Desconocido'}.")
