import logging
import asyncio
from concurrent.futures import ThreadPoolExecutor
from typing import Callable, Coroutine

from pymysqlreplication import BinLogStreamReader
from pymysqlreplication.row_event import (
    WriteRowsEvent,
    UpdateRowsEvent,
    DeleteRowsEvent,
)
from pymysqlreplication.event import QueryEvent, XidEvent

from .event_processor import EventProcessor

logger = logging.getLogger(__name__)

class BinlogReader:
    """
    Clase Productora Genérica.
    Lee el binlog de una base de datos de ORIGEN y pone las transacciones
    en una cola para un consumidor. Es reutilizable para cualquier flujo.
    """
    def __init__(self,
                 queue: asyncio.Queue,
                 connection_settings: dict,
                 resume_position: dict,
                 server_id: int,
                 included_dbs_str: str,
                 included_tables_str: str,
                 flow_name: str,
                 branch_id: int):
        
        self.queue = queue
        self.connection_settings = connection_settings
        self.resume_position = resume_position
        self.server_id = server_id
        self.included_dbs_str = included_dbs_str
        self.flow_name = flow_name
        self.branch_id = branch_id

        self.event_processor = EventProcessor(included_tables_str)
        logger.info(f"BinlogReader (Productor) para el flujo '{self.flow_name}' inicializado.")

    async def start(self):
        """
        Punto de entrada asíncrono. Ejecuta la lectura bloqueante en un hilo separado.
        """
        logger.info(f"-> [Productor - {self.flow_name}] Iniciando lectura de Binlog en un hilo...")
        loop = asyncio.get_running_loop()
        
        with ThreadPoolExecutor() as executor:
            await loop.run_in_executor(
                executor,
                self._run_blocking_reader
            )

    def _run_blocking_reader(self):
        """
        Lógica bloqueante de lectura del stream. Se ejecuta en un hilo.
        """
        log_file = self.resume_position['log_file']
        log_pos = self.resume_position['log_pos']
        
        logger.info(f"[{self.flow_name}] Conectando al stream del binlog desde: {log_file}:{log_pos}")

        stream = BinLogStreamReader(
            connection_settings=self.connection_settings,
            server_id=self.server_id,
            only_events=[WriteRowsEvent, UpdateRowsEvent, DeleteRowsEvent, QueryEvent, XidEvent],
            log_file=log_file,
            log_pos=log_pos,
            resume_stream=True,
            blocking=True,
            only_dbs=self.included_dbs_str.split(',') if self.included_dbs_str else None,
        )

        transaction_events = []
        is_in_transaction = False

        for binlogevent in stream:
            if isinstance(binlogevent, QueryEvent) and binlogevent.query.lower() == 'begin':
                is_in_transaction = True
                transaction_events = []
                continue

            # --- MANEJO DE AUTOCOMMIT ---
            # Si no estamos en una transacción y llega un evento de datos, es un autocommit.
            if not is_in_transaction and isinstance(binlogevent, (WriteRowsEvent, UpdateRowsEvent, DeleteRowsEvent)):
                processed_event = self.event_processor.process(binlogevent)
                if processed_event:
                    # Lo empaquetamos como una transacción de un solo evento
                    package = {
                        "source_id": self.branch_id,
                        "events": [processed_event],
                        "log_pos": {"log_file": stream.log_file, "log_pos": stream.log_pos}
                    }
                    future = asyncio.run_coroutine_threadsafe(self.queue.put(package), asyncio.get_running_loop())
                    future.result()
                    logger.info(f"-> [Productor - {self.flow_name}] Evento AUTOCOMMIT encolado.")
                continue
            # --- FIN MANEJO AUTOCOMMIT ---

            if is_in_transaction:
                processed_event = self.event_processor.process(binlogevent)
                if processed_event:
                    transaction_events.append(processed_event)

                if isinstance(binlogevent, XidEvent): # COMMIT
                    if transaction_events:
                        package = {
                            "source_id": self.branch_id,
                            "events": transaction_events,
                            "log_pos": {"log_file": stream.log_file, "log_pos": stream.log_pos}
                        }
                        future = asyncio.run_coroutine_threadsafe(self.queue.put(package), asyncio.get_running_loop())
                        future.result()
                        logger.info(f"-> [Productor - {self.flow_name}] Transacción encolada con {len(transaction_events)} eventos.")
                    
                    transaction_events = []
                    is_in_transaction = False
