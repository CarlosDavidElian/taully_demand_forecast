from __future__ import annotations

import json
from pathlib import Path

from docx import Document


ROOT = Path(__file__).resolve().parents[1]
DATA = json.loads((ROOT / ".tmp_anexo2" / "backtest_real_2026-05-19_a_2026-05-25.json").read_text(encoding="utf-8"))
FILES = [
    ("A", ROOT / "outputs" / "01a0b651-7620-73f0-8f6a-8b68b14f7a17" / "Anexo_2_Ficha_A_PMS_validacion_historica_7_dias.docx", "baseline_metrics"),
    ("B", ROOT / "outputs" / "01a0b651-7620-73f0-8f6a-8b68b14f7a17" / "Anexo_2_Ficha_B_modelo_ML_validacion_historica_7_dias.docx", "model_metrics"),
]


def result_text(value: float) -> str:
    return f"{value:.2f}".replace(".", ",")


for label, path, metric_key in FILES:
    document = Document(path)
    content = "\n".join(
        [paragraph.text for paragraph in document.paragraphs]
        + [cell.text for table in document.tables for row in table.rows for cell in row.cells]
    ).lower()
    expected = DATA["all_category_day_observations"][metric_key]
    observed = [row.cells[1].text for row in document.tables[2].rows[1:]]
    expected_values = [
        result_text(expected["mae"]),
        result_text(expected["rmse"]),
        result_text(expected["mape"]) + "%",
        result_text(expected["wape"]) + "%",
    ]
    if observed != expected_values:
        raise AssertionError(f"Ficha {label}: métricas distintas a la fuente: {observed} != {expected_values}")
    if len(document.tables[3].rows) != 6 or len(document.tables[4].rows) != 8:
        raise AssertionError(f"Ficha {label}: no conserva 5 categorías y 7 días.")
    if any(value in content for value in ("simulado", "16–22", "16-22", "n.a.", "may de", "april")):
        raise AssertionError(f"Ficha {label}: conserva texto de una versión no válida.")
    if "19 al 25 de mayo de 2026" not in content or "1 661 unidades" not in content:
        raise AssertionError(f"Ficha {label}: periodo o total real incorrecto.")
    print(f"OK ficha {label}: 6 tablas, 5 categorías, 7 días y métricas verificadas.")
