# api/routes.py
from fastapi import FastAPI
from core.state import shared_state

def create_app() -> FastAPI:
    """Fabrica y retorna la aplicacion FastAPI."""
    app = FastAPI(title="API de Monitoreo del Servicio de Replicación")

    @app.get("/status")
    async def get_status():
        """Devuelve el estado actual del servicio."""
        return shared_state.to_dict()

    return app
