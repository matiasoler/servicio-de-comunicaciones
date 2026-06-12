import logging
import asyncio
from typing import Callable, Coroutine
import aiomysql

logger = logging.getLogger(__name__)

class Transmitter:
    """
    Clase Consumidora Genérica.
    Su responsabilidad es tomar paquetes de transacciones de una cola,
    traducirlos a SQL y ejecutarlos en una base de datos de DESTINO.
    """
    def __init__(self,
                 queue: asyncio.Queue,
                 target_pool: aiomysql.Pool,
                 state_saver: Callable[[str, int], Coroutine],
                 flow_name: str):
        
        self.queue = queue
        self.target_pool = target_pool
        self.state_saver = state_saver
        self.flow_name = flow_name
        logger.info(f"Transmitter (Consumidor) para el flujo '{self.flow_name}' inicializado.")

    def _build_where_clause(self, row_data: dict) -> tuple[str, list]:
        """Construye una cláusula WHERE usando todos los valores de la fila."""
        clauses = []
        values = []
        for key, value in row_data.items():
            if value is None:
                clauses.append(f"`{key}` IS NULL")
            else:
                clauses.append(f"`{key}` = %s")
                values.append(value)
        if not clauses:
            raise ValueError("No se pueden construir cláusulas WHERE a partir de una fila vacía.")
        return " AND ".join(clauses), values

    async def start(self):
        """Punto de entrada asíncrono para iniciar el proceso de transmisión."""
        logger.info(f"<- [Consumidor - {self.flow_name}] Iniciando la espera de transacciones...")

        while True:
            transaction_package = await self.queue.get()
            logger.info(f"<- [Consumidor - {self.flow_name}] Nueva transacción con {len(transaction_package['events'])} eventos recibida.")
            
            async with self.target_pool.acquire() as conn:
                async with conn.cursor() as cursor:
                    try:
                        await conn.begin()
                        
                        for event in transaction_package['events']:
                            table = event['table']
                            event_type = event['type']
                            
                            if event_type == 'write':
                                for row in event['rows']:
                                    cols = ", ".join([f"`{k}`" for k in row.keys()])
                                    placeholders = ", ".join(["%s"] * len(row))
                                    sql = f"INSERT INTO `{table}` ({cols}) VALUES ({placeholders})"
                                    await cursor.execute(sql, list(row.values()))
                            
                            elif event_type == 'update':
                                for row in event['rows']:
                                    before_values = row['before_values']
                                    after_values = row['after_values']
                                    set_clause = ", ".join([f"`{k}` = %s" for k in after_values.keys()])
                                    where_clause, where_values = self._build_where_clause(before_values)
                                    sql = f"UPDATE `{table}` SET {set_clause} WHERE {where_clause}"
                                    await cursor.execute(sql, list(after_values.values()) + where_values)

                            elif event_type == 'delete':
                                for row in event['rows']:
                                    where_clause, where_values = self._build_where_clause(row['values'])
                                    sql = f"DELETE FROM `{table}` WHERE {where_clause}"
                                    await cursor.execute(sql, where_values)
                        
                        await conn.commit()
                        logger.info(f"[{self.flow_name}] Transacción confirmada en la base de datos de destino.")
                        
                        log_pos_info = transaction_package['log_pos']
                        await self.state_saver(log_pos_info['log_file'], log_pos_info['log_pos'])
                        logger.info(f"<- [Consumidor - {self.flow_name}] Nueva posición guardada: {log_pos_info['log_file']}:{log_pos_info['log_pos']}")

                    except Exception as e:
                        logger.error(f"[{self.flow_name}] Error al procesar la transacción en la BD de destino: {e}. Se hará ROLLBACK.", exc_info=True)
                        await conn.rollback()
                        await asyncio.sleep(10)
                    finally:
                        self.queue.task_done()
