# core/state.py
from dataclasses import dataclass, field
from typing import List

@dataclass
class ServiceState:
    my_branch_id: int | None = None
    central_branch_id: int | None = None
    outbound_status: str = "iniciando"
    inbound_status: str = "iniciando"
    last_outbound_sync: str | None = None
    last_inbound_sync: str | None = None
    errors: List[str] = field(default_factory=list)

    def to_dict(self):
        return {
            "my_branch_id": self.my_branch_id,
            "central_branch_id": self.central_branch_id,
            "outbound_status": self.outbound_status,
            "inbound_status": self.inbound_status,
            "last_outbound_sync": self.last_outbound_sync,
            "last_inbound_sync": self.last_inbound_sync,
            "errors": self.errors[-10:] # Mostrar solo los ultimos 10 errores
        }

# Instancia global del estado
shared_state = ServiceState()
