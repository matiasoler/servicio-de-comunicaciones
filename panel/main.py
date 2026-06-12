# Archivo: panel_orquestador/main.py
from fastapi import FastAPI, Request
import datetime

app = FastAPI(title="Orquestador V&S")

# Endpoint para recibir las ventas/eventos del Agente
@app.post("/api/v1/sincronizar")
async def recibir_evento(request: Request):
    # Tomamos el JSON que nos manda el agente
    datos = await request.json()
    hora_actual = datetime.datetime.now().strftime("%H:%M:%S")
    
    print(f"[{hora_actual}] 📥 DATOS RECIBIDOS de Sucursal {datos.get('id_sucursal')}:")
    print(f"   -> Tabla: {datos.get('tabla')} | Acción: {datos.get('accion')}")
    print(f"   -> Payload: {datos.get('datos')}")
    print("-" * 40)
    
    return {"status": "ok", "mensaje": "Evento registrado en BD Central"}

# Endpoint para que el Agente consulte si hay comandos
@app.get("/api/v1/comandos/{id_sucursal}")
async def consultar_comandos(id_sucursal: int):
    # Por ahora simulamos que no hay comandos críticos
    return {"comando": "NINGUNO"}