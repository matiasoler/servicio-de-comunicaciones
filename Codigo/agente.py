import asyncio
import logging

from src.config import Config
from src.database import DatabaseManager
from src.binlog_reader import BinlogReader
from src.transmitter import Transmitter

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger('agente')

async def main():
    """
    Orquestador principal del agente bidireccional.
    Crea, configura y lanza los flujos de subida y bajada.
    """
    logger.info("Iniciando orquestador del agente bidireccional...")
    db_manager = None
    
    try:
        # 1. Cargar configuración estática del .ini
        config = Config()

        # 2. Inicializar el manejador de las bases de datos
        db_manager = DatabaseManager(config)
        await db_manager.initialize()

        # 3. Crear colas de comunicación para cada flujo
        upload_queue = asyncio.Queue(maxsize=config.getint('replication', 'max_queue_size', fallback=100))
        download_queue = asyncio.Queue(maxsize=config.getint('replication', 'max_queue_size', fallback=100))

        # --- CONFIGURACIÓN DEL FLUJO DE SUBIDA (LOCAL -> CENTRAL) ---
        upload_reader = BinlogReader(
            queue=upload_queue,
            connection_settings=config['database_local'],
            resume_position=db_manager.get_upload_resume_position(),
            server_id=config.getint('replication', 'local_server_id'),
            included_dbs_str=config.get('replication', 'included_dbs'),
            included_tables_str=config.get('replication', 'included_tables'),
            flow_name="Subida",
            branch_id=config.getint('general', 'branch_id')
        )
        upload_transmitter = Transmitter(
            queue=upload_queue,
            target_pool=db_manager.get_central_pool(),
            state_saver=db_manager.save_upload_position,
            flow_name="Subida"
        )

        # --- CONFIGURACIÓN DEL FLUJO DE BAJADA (CENTRAL -> LOCAL) ---
        # Es crucial usar un server_id diferente para cada conexión de replicación a una misma BD.
        central_server_id = config.getint('replication', 'central_server_id')
        download_reader = BinlogReader(
            queue=download_queue,
            connection_settings=db_manager.dynamic_config['database_central'],
            resume_position=db_manager.get_download_resume_position(),
            server_id=central_server_id,
            included_dbs_str=config.get('replication', 'included_dbs'),
            included_tables_str=config.get('replication', 'included_tables'),
            flow_name="Bajada",
            branch_id=config.getint('general', 'branch_id')
        )
        download_transmitter = Transmitter(
            queue=download_queue,
            target_pool=db_manager.get_local_pool(),
            state_saver=db_manager.save_download_position,
            flow_name="Bajada"
        )
        
        # 5. Lanzar todas las tareas y mantenerlas corriendo
        logger.info("Lanzando flujos de Subida y Bajada en paralelo...")
        tasks = [
            asyncio.create_task(upload_reader.start()),
            asyncio.create_task(upload_transmitter.start()),
            asyncio.create_task(download_reader.start()),
            asyncio.create_task(download_transmitter.start())
        ]
        await asyncio.gather(*tasks)

    except FileNotFoundError as e:
        logger.error(f"Error de configuración: {e}", exc_info=False)
    except Exception as e:
        logger.error(f"Error crítico en el `main loop`: {e}", exc_info=True)
    finally:
        logger.info("Cerrando recursos...")
        if db_manager:
            await db_manager.close()

if __name__ == "__main__":
    logger.info("Iniciando Agente de Comunicaciones Bidireccional...")
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logger.info("Agente detenido manualmente (Ctrl+C).")
    finally:
        logger.info("Proceso del agente finalizado.")
