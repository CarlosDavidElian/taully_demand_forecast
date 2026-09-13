"""Funciones compartidas para comprobar que un modelo corresponde al historial."""

from __future__ import annotations

import hashlib
from typing import Iterable

from domain.entities.demand import Demand


def history_fingerprint(demands: Iterable[Demand]) -> str:
    """Crea una firma estable del contenido, no de la fecha del archivo.

    Así un pronóstico se bloquea cuando se agregan o corrigen ventas, incluso
    si la modificación del archivo conserva una fecha de venta antigua.
    """
    rows = sorted(
        (
            demand.date.strftime("%Y-%m-%d"),
            str(demand.category).strip(),
            f"{float(demand.quantity):.10f}",
        )
        for demand in demands
    )
    digest = hashlib.sha256()
    for row in rows:
        digest.update("|".join(row).encode("utf-8"))
        digest.update(b"\n")
    return digest.hexdigest()
