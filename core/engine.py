# core/engine.py
import asyncio
import datetime
import aiomysql
from pymysqlreplication import BinLogStreamReader
from pymysqlreplication.row_event import (
    DeleteRowsEvent,
    UpdateRowsEvent,
    WriteRowsEvent,
)
from pymysqlreplication.event import QueryEvent, XidEvent
from core.state import shared_state
from core.buffer import TransactionBuffer, TransactionBlock

class ReplicationEngine:
    def __init__(self, name: str, source_pool: aiomysql.Pool, dest_pool: aiomysql.Pool, excepted_tables: set, state_key: str, get_start_pos_callback, update_logs_callback, server_id: int, max_blocks_per_batch: int = 10):
        self.name = name
        self.source_pool = source_pool
        self.dest_pool = dest_pool
        self.excepted_tables = excepted_tables
        self.state_key = state_key
        self.get_start_pos_callback = get_start_pos_callback
        self.update_logs_callback = update_logs_callback
        self.server_id = server_id
        self.max_blocks_per_batch = max_blocks_per_batch
        self.running = True

    def _build_mysql_settings_from_pool(self, pool: aiomysql.Pool) -> dict:
        """Extrae la configuración de conexión del pool para pymysqlreplication."""
        return {
            "host": pool._host,
            "port": pool._port,
            "user": pool._user,
            "passwd": pool._password,
        }

    async def run(self):
        print(f"[{self.name}] Trabajador iniciado.")
        setattr(shared_state, self.state_key, "corriendo")
        
        while self.running:
            try:
                log_to_read = await self.get_start_pos_callback(self.source_pool, self.dest_pool)

                if not log_to_read:
                    print(f"[{self.name}] Esperando condiciones para iniciar lectura de logs...")
                    await asyncio.sleep(10)
                    continue
                
                mysql_settings = self._build_mysql_settings_from_pool(self.source_pool)
                
                print(f"[{self.name}] Conectando al binlog desde {log_to_read['archivo']}:{log_to_read['pos']}...")
                
                # Leemos eventos DML y eventos de Query/Xid para armar los bloques (BEGIN/COMMIT)
                stream = BinLogStreamReader(
                    connection_settings=mysql_settings,
                    server_id=self.server_id,
                    log_file=log_to_read['archivo'],
                    log_pos=log_to_read['pos'],
                    resume_stream=True,
                    only_events=[DeleteRowsEvent, WriteRowsEvent, UpdateRowsEvent, QueryEvent, XidEvent]
                )

                loop = asyncio.get_running_loop()
                buffer = TransactionBuffer(max_blocks_per_batch=self.max_blocks_per_batch)

                while self.running:
                    # Lectura asíncrona (el recepcionista esperando el paquete)
                    binlogevent = await loop.run_in_executor(None, stream.fetchone)
                    
                    if binlogevent is None:
                        break
                    if not self.running:
                        break

                    # 1. Alimentar el Buffer con el nuevo evento
                    buffer.process_event(binlogevent, stream.log_file, stream.log_pos, self.excepted_tables)

                    # 2. Revisar el "Grifo": ¿Hay bloques completos listos para enviar?
                    if buffer.has_ready_blocks():
                        batch = buffer.get_batch()
                        
                        for block in batch:
                            try:
                                # Aplicar la transacción completa
                                await self._apply_block_to_dest(block)
                                
                                # Si no hubo error, confirmamos avanzando los logs
                                await self.update_logs_callback(
                                    block.end_log_pos, block.end_log_file,
                                    self.source_pool, self.dest_pool
                                )
                                
                                print(f"[{self.name}] Lote aplicado OK. Nueva pos: {block.end_log_file}:{block.end_log_pos}")
                                setattr(shared_state, f"last_{self.name}_sync", datetime.datetime.now().isoformat())
                            
                            except Exception as e_apply:
                                print(f"[{self.name}] Error aplicando lote: {e_apply}. Deteniendo stream para reintentar.")
                                stream.close()
                                buffer.clear() # Limpiamos el buffer por seguridad
                                raise e_apply 

                stream.close()

            except Exception as e:
                error_msg = f"[{self.name}] Error general: {e}"
                print(error_msg)
                shared_state.errors.append(error_msg)
                setattr(shared_state, self.state_key, f"error: {e}")
                await asyncio.sleep(30)

            await asyncio.sleep(5)

    async def _apply_block_to_dest(self, block: TransactionBlock):
        """Ejecuta todos los eventos de un bloque dentro de una única transacción."""
        async with self.dest_pool.acquire() as conn:
            # Empezamos LA TRANSACCION en la base de datos destino
            await conn.begin()
            try:
                async with conn.cursor() as cursor:
                    for event in block.events:
                        table = event.table
                        
                        if isinstance(event, WriteRowsEvent):
                            for row in event.rows:
                                data = row["values"]
                                columns = ", ".join(data.keys())
                                placeholders = ", ".join(["%s"] * len(data))
                                sql = f"INSERT INTO {table} ({columns}) VALUES ({placeholders})"
                                values = tuple(data.values())
                                print(f"      -> Ejecutando: {sql}")
                                await cursor.execute(sql, values)
                                
                        elif isinstance(event, UpdateRowsEvent):
                            for row in event.rows:
                                before = row["before_values"]
                                after = row["after_values"]
                                set_clause = ", ".join([f"{k} = %s" for k in after.keys()])
                                where_clause = " AND ".join([f"{k} = %s" for k in before.keys()])
                                sql = f"UPDATE {table} SET {set_clause} WHERE {where_clause}"
                                values = tuple(after.values()) + tuple(before.values())
                                print(f"      -> Ejecutando: {sql}")
                                await cursor.execute(sql, values)
                                
                        elif isinstance(event, DeleteRowsEvent):
                            for row in event.rows:
                                data = row["values"]
                                where_clause = " AND ".join([f"{k} = %s" for k in data.keys()])
                                sql = f"DELETE FROM {table} WHERE {where_clause}"
                                values = tuple(data.values())
                                print(f"      -> Ejecutando: {sql}")
                                await cursor.execute(sql, values)

                # Si todos los eventos del bloque pasaron bien, confirmamos (COMMIT)
                await conn.commit()
                print(f"      [✓] Transacción (Bloque de {len(block.events)} eventos) aplicada con éxito.")

            except aiomysql.IntegrityError as e:
                # Si un solo evento falla, deshacemos TODO el bloque (ROLLBACK)
                await conn.rollback()
                # TODO (Future Check): Consultar information_schema.KEY_COLUMN_USAGE
                print(f"      [!] Error de Integridad (FK) en transacción: {e}. ROLLBACK ejecutado.")
                raise e
            except Exception as e:
                await conn.rollback()
                print(f"      [!] Error aplicando transacción: {e}. ROLLBACK ejecutado.")
                raise e

    def stop(self):
        self.running = False
        print(f"[{self.name}] Solicitud de detención recibida.")
        self.running = False
        print(f"[{self.name}] Solicitud de detención recibida.")
