import logging
from pymysqlreplication.row_event import WriteRowsEvent, UpdateRowsEvent, DeleteRowsEvent

logger = logging.getLogger(__name__)

class EventProcessor:
    """
    Procesa eventos crudos del binlog y los transforma en un formato
    serializable y estandarizado, filtrando solo las tablas que nos interesan.
    """
    def __init__(self, included_tables_str: str):
        if not included_tables_str:
            logger.warning("No se especificaron 'included_tables'. Se procesarán todas las tablas.")
            self.included_tables = None
        else:
            self.included_tables = set(table.strip().lower() for table in included_tables_str.split(','))
            logger.info(f"EventProcessor filtrará por las siguientes tablas: {self.included_tables}")

    def process(self, event):
        """
        Filtra y transforma un evento del binlog.
        Devuelve un diccionario si el evento es de una tabla de interés, o None si no lo es.
        """
        table_name = event.table.lower()

        # Si tenemos una lista de tablas y el evento no es de una de ellas, lo ignoramos.
        if self.included_tables and table_name not in self.included_tables:
            return None

        event_type = None
        rows = []

        if isinstance(event, WriteRowsEvent):
            event_type = 'write'
            # 'values' es una lista de diccionarios, donde cada diccionario es una fila.
            rows = event.rows
        
        elif isinstance(event, UpdateRowsEvent):
            event_type = 'update'
            # Para updates, cada "row" es un diccionario con 'before_values' y 'after_values'.
            # Nos quedamos solo con los valores después de la actualización.
            rows = [row['after_values'] for row in event.rows]

        elif isinstance(event, DeleteRowsEvent):
            event_type = 'delete'
            # Para deletes, 'values' contiene los valores de las filas eliminadas.
            rows = event.rows
        
        # Si no es un evento de datos que nos interese, lo ignoramos.
        else:
            return None
        
        logger.debug(f"Evento procesado: tabla '{table_name}', tipo '{event_type}', {len(rows)} filas.")

        return {
            "table": table_name,
            "schema": event.schema,
            "type": event_type,
            "rows": rows
        }

