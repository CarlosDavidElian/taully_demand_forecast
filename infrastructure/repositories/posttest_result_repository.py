"""Almacena resultados de postest como evidencia reproducible en JSON."""

from __future__ import annotations

import json
import os
import re
import uuid
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


class PosttestResultRepository:
    """Guarda cada ejecución sin reemplazar los resultados anteriores."""

    _RUN_ID_PATTERN = re.compile(r"^[a-z0-9_-]{12,80}$")

    def __init__(self, directory: Path):
        self.directory = Path(directory)

    def save(self, result: dict[str, Any]) -> dict[str, Any]:
        self.directory.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.now(timezone.utc)
        cutoff = str(result.get("cutoff_date") or "sin_corte")
        horizon = int(result.get("horizon_days") or 0)
        run_id = (
            f"postest_{timestamp.strftime('%Y%m%dT%H%M%SZ')}_"
            f"{cutoff.replace('-', '')}_{horizon}d_{uuid.uuid4().hex[:8]}"
        ).lower()
        record = {
            "run_id": run_id,
            "created_at": timestamp.isoformat(),
            "result": result,
        }
        target = self.directory / f"{run_id}.json"
        temporary = self.directory / f".{run_id}.{uuid.uuid4().hex}.tmp"
        temporary.write_text(
            json.dumps(record, ensure_ascii=False, allow_nan=False, indent=2),
            encoding="utf-8",
        )
        os.replace(temporary, target)
        return record

    def get(self, run_id: str) -> dict[str, Any]:
        normalized = str(run_id or "").strip().lower()
        if not self._RUN_ID_PATTERN.fullmatch(normalized):
            raise ValueError("El identificador del postest no es válido.")
        path = self.directory / f"{normalized}.json"
        if not path.is_file():
            raise FileNotFoundError("No se encontró el resultado del postest solicitado.")
        try:
            record = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError as exc:
            raise ValueError("El resultado guardado del postest no se puede leer.") from exc
        if record.get("run_id") != normalized or not isinstance(record.get("result"), dict):
            raise ValueError("El resultado guardado del postest no tiene una estructura válida.")
        return record
