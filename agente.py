# Wrote agente.py
import asyncio
import logging
# Configuración básica de logging para ver qué está pasando.
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(threadName)s - %(levelname)s - %(message)s')
async def lector_binlog(queue: asyncio.Queue):
    """
    Productor (Producer):
    Su única responsabilidad es leer eventos del binlog de MariaDB
    y ponerlos en una cola (queue) para ser procesados.
    """
    logging.info("-> [Productor] Iniciando lector de binlog...")
    
    # --- A IMPLEMENTAR ---
    # Aquí irá la lógica de conexión a MariaDB con `python-mysql-replication`.
    # IMPORTANTE: Esa librería es bloqueante, por lo que tendremos que ejecutarla
    # en un `ThreadPoolExecutor` para no frenar el event loop de asyncio.
    # loop.run_in_executor(executor, sync_blocking_function, *args)
    # ---------------------
    # Simulación de lectura de eventos para probar la estructura.
    event_count = 0
    while True:
        await asyncio.sleep(2) # Simula la espera de nuevos eventos en el binlog.
        
        # En la realidad, este evento sería un objeto con los datos de la fila/transacción.
        event = {"id": event_count, "data": f"datos de la venta {event_count}"}
        
        logging.info(f"-> [Productor] Nuevo evento detectado: {event}")
        await queue.put(event)
        event_count += 1
async def transmisor(queue: asyncio.Queue):
    """
    Consumidor (Consumer):
    Toma eventos de la cola, los procesa (si es necesario) y los transmite
    a la base de datos central.
    """
    logging.info("<- [Consumidor] Iniciando transmisor de datos...")
    
    while True:
        # Espera de forma asíncrona hasta que haya un evento disponible en la cola.
        event = await queue.get()
        
        logging.info(f"<- [Consumidor] Procesando evento: {event}")
        
        # --- A IMPLEMENTAR ---
        # Aquí irá la lógica para:
        # 1. Agrupar eventos en una transacción si es necesario.
        # 2. Conectarse a la base de datos central.
        # 3. Insertar los datos.
        # 4. Manejar errores y reintentos.
        # ---------------------
        # Simulación de una operación de red/base de datos que toma tiempo.
        await asyncio.sleep(0.5) 
        
        logging.info(f"<- [Consumidor] Evento {event['id']} transmitido con éxito.")
        
        # Le avisa a la cola que terminamos de procesar este item.
        # Es importante para la gestión de la cola.
        queue.task_done()
async def main():
    """
    Función principal que orquesta el Productor y el Consumidor.
    """
    # Creamos la cola de comunicación. `maxsize` es importante para aplicar
    # "backpressure": si el consumidor es lento, el productor esperará
    # en lugar de llenar la memoria con eventos pendientes.
    queue = asyncio.Queue(maxsize=100)
    # Creamos las tareas para el productor y el consumidor.
    productor_task = asyncio.create_task(lector_binlog(queue))
    consumidor_task = asyncio.create_task(transmisor(queue))
    # Esperamos a que ambas tareas finalicen (en este caso, nunca, hasta que se cancele).
    await asyncio.gather(productor_task, consumidor_task)
if __name__ == "__main__":
    logging.info("Iniciando Agente de Comunicaciones...")
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        logging.info("Agente detenido manualmente (Ctrl+C).")