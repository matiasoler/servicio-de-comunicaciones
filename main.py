# main.py
import asyncio
import aiomysql
import uvicorn
from core.state import shared_state
from core.engine import ReplicationEngine
from api.routes import create_app
from database.connection import (
    get_local_db_config, create_db_pool, 
    discover_identity, discover_central, get_excepted_tables,
    verify_binlog_config
)

# --- Callbacks para inyectar lógica asimétrica en el motor ---

async def get_ida_start_pos(source_pool, dest_pool):
    """Local: sucursales_log_envio (my_id)"""
    async with source_pool.acquire() as conn:
        async with conn.cursor(aiomysql.DictCursor) as cursor:
            await cursor.execute("SELECT archivo_log as archivo, posicion as pos FROM sucursales_log_envio WHERE id_sucursal = %s", (shared_state.my_branch_id,))
            return await cursor.fetchone()

async def update_ida_logs(new_pos, new_file, source_pool, dest_pool):
    """Actualiza envio local (my_id) y recepcion central (my_id)"""
    async with source_pool.acquire() as conn:
        async with conn.cursor() as cursor:
            await cursor.execute(
                "UPDATE sucursales_log_envio SET posicion = %s, archivo_log = %s, fecha_actualizacion = NOW() WHERE id_sucursal = %s",
                (new_pos, new_file, shared_state.my_branch_id)
            )
    async with dest_pool.acquire() as conn:
        async with conn.cursor() as cursor:
            await cursor.execute(
                "UPDATE sucursales_log_recepcion SET posicion = %s, archivo_log = %s, fecha_actualizacion = NOW() WHERE id_sucursal = %s",
                (new_pos, new_file, shared_state.my_branch_id)
            )

async def get_vuelta_start_pos(source_pool, dest_pool):
    """Central: sucursales_log_envio (my_id) vs Local: sucursales_log_recepcion (central_id)"""
    async with source_pool.acquire() as conn_central:
        async with conn_central.cursor(aiomysql.DictCursor) as cursor_central:
            await cursor_central.execute("SELECT archivo_log as archivo, posicion as pos FROM sucursales_log_envio WHERE id_sucursal = %s", (shared_state.my_branch_id,))
            envio_central = await cursor_central.fetchone()
            
    async with dest_pool.acquire() as conn_local:
        async with conn_local.cursor(aiomysql.DictCursor) as cursor_local:
            await cursor_local.execute("SELECT archivo_log as archivo, posicion as pos FROM sucursales_log_recepcion WHERE id_sucursal = %s", (shared_state.central_branch_id,))
            recepcion_local = await cursor_local.fetchone()

    if not envio_central or not recepcion_local:
        return None
    if envio_central['archivo'] != recepcion_local['archivo'] or envio_central['pos'] != recepcion_local['pos']:
        return None
    return envio_central

async def update_vuelta_logs(new_pos, new_file, source_pool, dest_pool):
    """Actualiza envio central (my_id) y recepcion local (central_id)"""
    async with source_pool.acquire() as conn_central:
        async with conn_central.cursor() as cursor_central:
            await cursor_central.execute(
                "UPDATE sucursales_log_envio SET posicion = %s, archivo_log = %s, fecha_actualizacion = NOW() WHERE id_sucursal = %s",
                (new_pos, new_file, shared_state.my_branch_id)
            )
    async with dest_pool.acquire() as conn_local:
        async with conn_local.cursor() as cursor_local:
            await cursor_local.execute(
                "UPDATE sucursales_log_recepcion SET posicion = %s, archivo_log = %s, fecha_actualizacion = NOW() WHERE id_sucursal = %s",
                (new_pos, new_file, shared_state.central_branch_id)
            )

# --- Fin Callbacks ---

async def main():
    print("--- Iniciando Servicio de Replicación (Arquitectura Modular) ---")
    
    # 1. Conexión Local y Descubrimiento
    local_config = get_local_db_config('Codigo/config.ini')
    local_pool = await create_db_pool(local_config)
    
    try:
        await verify_binlog_config(local_pool, "LOCAL")
        print("-> Check: Binlog local activado y en formato ROW.")

        my_id = await discover_identity(local_pool)
        shared_state.my_branch_id = my_id
        print(f"-> Identidad local confirmada: id_sucursal={my_id}")

        central_id, central_config = await discover_central(local_pool)
        shared_state.central_branch_id = central_id
        print(f"-> Central encontrada: id_sucursal={central_id}")
        
        outbound_exceptions = await get_excepted_tables(local_pool, central_id)
        print(f"-> Tablas exceptuadas (envío a central): {len(outbound_exceptions)} reglas.")
    except Exception as e:
        print(f"Error en inicialización local: {e}")
        local_pool.close()
        await local_pool.wait_closed()
        return

    # 2. Conexión Central y Descubrimiento
    try:
        central_pool = await create_db_pool(central_config)
        
        await verify_binlog_config(central_pool, "CENTRAL")
        print("-> Check: Binlog central activado y en formato ROW.")

        inbound_exceptions = await get_excepted_tables(central_pool, my_id)
        print(f"-> Tablas exceptuadas (recepción desde central): {len(inbound_exceptions)} reglas.")
    except Exception as e:
        print(f"Error conectando a central: {e}")
        # En el futuro, podríamos querer que el agente corra igual, aunque la central esté caída,
        # encolando trabajo. Por ahora, abortamos si no hay central al arrancar.
        local_pool.close()
        await local_pool.wait_closed()
        return

    # 3. Levantar los Motores (Trabajadores)
    ida_engine = ReplicationEngine(
        name="outbound", 
        source_pool=local_pool, 
        dest_pool=central_pool, 
        excepted_tables=outbound_exceptions, 
        state_key="outbound_status",
        get_start_pos_callback=get_ida_start_pos,
        update_logs_callback=update_ida_logs,
        server_id=my_id + 1000
    )
    
    vuelta_engine = ReplicationEngine(
        name="inbound", 
        source_pool=central_pool, 
        dest_pool=local_pool, 
        excepted_tables=inbound_exceptions, 
        state_key="inbound_status",
        get_start_pos_callback=get_vuelta_start_pos,
        update_logs_callback=update_vuelta_logs,
        server_id=central_id + 2000
    )

    # 4. Levantar la API de Monitoreo
    app = create_app()
    config_uvicorn = uvicorn.Config(app, host="0.0.0.0", port=8000, log_level="warning")
    server = uvicorn.Server(config_uvicorn)

    # 5. Ejecutar Todo en Paralelo
    print("--- Todas las validaciones OK. Lanzando procesos en paralelo ---")
    await asyncio.gather(
        ida_engine.run_ida_loop(),
        vuelta_engine.run_vuelta_loop(),
        server.serve()
    )

if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        print("\nServicio detenido por el usuario (Ctrl+C).")
    except Exception as e:
        print(f"Error fatal: {e}")
