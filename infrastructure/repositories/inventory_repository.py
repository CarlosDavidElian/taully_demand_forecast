"""Almacenamiento inmutable para archivos de inventario importados.

El pronóstico no debe modificar el archivo que entrega la tienda.  Este
repositorio conserva una copia identificada por el contenido del archivo, de
modo que dos cargas idénticas reutilizan la misma copia y una nueva carga no
sobrescribe las anteriores.
"""

from __future__ import annotations

from dataclasses import dataclass
import hashlib
import os
from pathlib import Path
import shutil
import tempfile


@dataclass(frozen=True)
class StoredInventoryFile:
    """Referencia a una copia inmutable de un inventario importado."""

    inventory_id: str
    path: Path
    original_name: str


class InventoryWorkbookStore:
    """Guarda cada Excel importado sin reemplazar importaciones previas."""

    def __init__(self, storage_directory: str | Path):
        self.storage_directory = Path(storage_directory)

    def persist(self, source_path: str | Path) -> StoredInventoryFile:
        source = Path(source_path)
        if not source.exists() or not source.is_file():
            raise FileNotFoundError("No se encontró el archivo de inventario seleccionado.")

        digest = self._sha256(source)
        suffix = source.suffix.lower() or ".xlsx"
        target = self.storage_directory / f"{digest}{suffix}"
        self.storage_directory.mkdir(parents=True, exist_ok=True)

        # Un mismo contenido ya tiene una copia segura. No se reescribe para
        # conservar su fecha de importación y evitar pérdida de evidencias.
        if target.exists():
            if self._sha256(target) != digest:
                raise ValueError("La copia guardada de inventario no coincide con el archivo importado.")
            return StoredInventoryFile(digest, target, source.name)

        temporary_path: Path | None = None
        try:
            with tempfile.NamedTemporaryFile(
                suffix=".tmp",
                prefix=f".{digest[:12]}.",
                dir=self.storage_directory,
                delete=False,
            ) as temporary_file:
                temporary_path = Path(temporary_file.name)
            shutil.copyfile(source, temporary_path)
            # Un enlace duro se crea de forma atómica y falla si otra petición
            # ya guardó ese contenido. A diferencia de ``os.replace``, nunca
            # reemplaza una copia de inventario que ya existe.
            try:
                os.link(temporary_path, target)
            except FileExistsError:
                if self._sha256(target) != digest:
                    raise ValueError("No se pudo conservar una copia coherente del inventario.")
        finally:
            if temporary_path is not None and temporary_path.exists():
                temporary_path.unlink(missing_ok=True)

        return StoredInventoryFile(digest, target, source.name)

    @staticmethod
    def _sha256(path: Path) -> str:
        digest = hashlib.sha256()
        with path.open("rb") as source:
            for block in iter(lambda: source.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
