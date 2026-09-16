from dataclasses import dataclass
from datetime import datetime

@dataclass
class Demand:
    """Representa la demanda consolidada de una categoría comercial en un día."""
    date: datetime
    category: str
    quantity: float

    def __post_init__(self):
        if self.quantity < 0:
            raise ValueError("La demanda no puede ser negativa")
