# core/buffer.py
from pymysqlreplication.row_event import WriteRowsEvent, UpdateRowsEvent, DeleteRowsEvent
from pymysqlreplication.event import QueryEvent, XidEvent

class TransactionBlock:
    """
    Representa un bloque atómico de eventos que deben aplicarse juntos.
    """
    def __init__(self):
        self.events = []
        self.end_log_file = None
        self.end_log_pos = None

    def is_empty(self):
        return len(self.events) == 0


class TransactionBuffer:
    """
    Agrupa eventos individuales del binlog en bloques transaccionales.
    """
    def __init__(self, max_blocks_per_batch: int = 10):
        self.max_blocks_per_batch = max_blocks_per_batch
        self.ready_blocks = []      # Lista de TransactionBlock listos para enviar
        self.current_block = None   # El bloque que se está construyendo
        self.in_transaction = False

    def process_event(self, event, current_file: str, current_pos: int, excepted_tables: set):
        """
        Recibe un evento del binlog, lo evalúa y lo coloca en el buffer.
        """
        # 1. Manejo de transacciones explícitas (BEGIN)
        if isinstance(event, QueryEvent) and event.query.strip().upper() == "BEGIN":
            if self.in_transaction:
                # Ya estábamos en una transacción, esto es una anomalía o un anidamiento no soportado
                raise Exception("Error de consistencia: Se recibió un BEGIN pero ya hay una transacción abierta.")
            
            self.in_transaction = True
            self.current_block = TransactionBlock()
            return

        # 2. Manejo de fin de transacciones (COMMIT o XID)
        # XidEvent ocurre en InnoDB cuando se hace COMMIT. 
        # A veces puede venir un QueryEvent con "COMMIT".
        is_commit = isinstance(event, XidEvent) or (isinstance(event, QueryEvent) and event.query.strip().upper() == "COMMIT")
        if is_commit:
            if not self.in_transaction:
                # Recibimos un COMMIT suelto sin un BEGIN previo. Lo ignoramos o logueamos.
                # Es normal en algunas operaciones autocommit, no es crítico.
                return
            
            # Cerramos el bloque actual
            self.in_transaction = False
            if self.current_block and not self.current_block.is_empty():
                # Le estampamos la posición final donde ocurrió el COMMIT
                self.current_block.end_log_file = current_file
                self.current_block.end_log_pos = current_pos
                self.ready_blocks.append(self.current_block)
            self.current_block = None
            return

        # 3. Manejo de eventos de datos (DML)
        if isinstance(event, (WriteRowsEvent, UpdateRowsEvent, DeleteRowsEvent)):
            # Validar excepciones antes de agregar al bloque
            if event.table in excepted_tables:
                return # Ignoramos este evento

            if self.in_transaction:
                # Agregamos al bloque que se está armando
                if self.current_block is None:
                    self.current_block = TransactionBlock()
                self.current_block.events.append(event)
            else:
                # Operación autocommit (fuera de un BEGIN). Creamos un bloque único.
                single_block = TransactionBlock()
                single_block.events.append(event)
                single_block.end_log_file = current_file
                single_block.end_log_pos = current_pos
                self.ready_blocks.append(single_block)

    def has_ready_blocks(self) -> bool:
        """Indica si hay bloques terminados esperando ser enviados."""
        return len(self.ready_blocks) > 0

    def get_batch(self) -> list:
        """
        Retorna un lote de bloques (hasta el límite configurado) y los remueve del buffer.
        """
        if not self.ready_blocks:
            return []
        
        # Tomamos hasta max_blocks_per_batch
        batch_size = min(len(self.ready_blocks), self.max_blocks_per_batch)
        batch = self.ready_blocks[:batch_size]
        
        # Removemos los que tomamos
        self.ready_blocks = self.ready_blocks[batch_size:]
        
        return batch

    def clear(self):
        """Limpia el buffer (útil si hay un error fatal y hay que resetear el estado)."""
        self.ready_blocks.clear()
        self.current_block = None
        self.in_transaction = False
