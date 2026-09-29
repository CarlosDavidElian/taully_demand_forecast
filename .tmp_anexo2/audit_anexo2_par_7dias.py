"""Auditoría estructural y numérica del par pretest/postest simulado de 7 días."""

from __future__ import annotations

import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from docx import Document

ROOT = Path(r"C:\taully_demand_forecast")
sys.path.insert(0, str(ROOT / ".tmp_anexo2"))

from build_anexo2_posttest_prueba import generate_test_event, output_path, scenario_label


def clean(value: object) -> str:
    return " ".join(str(value or "").split())


def number(value: object) -> float:
    return float(str(value).replace(",", "").replace("%", ""))


def rows(table):
    return [[clean(cell.text) for cell in row.cells] for row in table.rows][1:]


def fail(message: str) -> None:
    raise AssertionError(message)


def audit_document(scenario: str):
    label = scenario_label(scenario)
    expected_categories, expected_products = generate_test_event(scenario)
    docx_path = output_path(scenario)
    if not docx_path.exists():
        fail(f"No existe el archivo {docx_path.name}.")
    document = Document(docx_path)
    headings = [clean(paragraph.text) for paragraph in document.paragraphs if clean(paragraph.text).startswith("Instrumento ")]
    if len(headings) != 8 or len(document.tables) != 17:
        fail(f"{label}: no tiene las ocho fichas y sus 17 tablas.")
    document_text = " ".join(
        clean(paragraph.text) for paragraph in document.paragraphs
    ) + " " + " ".join(clean(cell.text) for table in document.tables for row in table.rows for cell in row.cells)
    for required in (f"{label} simulado", "mismo período", "105 observaciones", "No representa ventas"):
        if required.lower() not in document_text.lower():
            fail(f"{label}: falta el aviso obligatorio '{required}'.")
    with zipfile.ZipFile(docx_path) as archive:
        xml = ET.fromstring(archive.read("word/document.xml"))
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    if len(xml.findall(".//w:br[@w:type='page']", ns)) != 8:
        fail(f"{label}: no conserva los ocho saltos de página.")

    expected_by_product = {item["product"]: item for item in expected_products}
    if len(expected_by_product) != 15:
        fail(f"{label}: la fuente no contiene 15 productos distintos.")
    balances = {}
    for row in rows(document.tables[2]):
        product, si, en, sf, cd = row[1], number(row[3]), number(row[4]), number(row[5]), number(row[6])
        expected = expected_by_product.get(product)
        if expected is None or cd != si + en - sf:
            fail(f"{label}: CD no cuadra en {product}.")
        if (si, en, sf, cd) != (expected["si"], expected["en"], expected["sf"], expected["cd"]):
            fail(f"{label}: el balance no coincide con la fuente en {product}.")
        if min(si, en, sf, cd) < 0:
            fail(f"{label}: hay una cantidad negativa en {product}.")
        balances[product] = (si, en, sf, cd)
    if len(balances) != 15:
        fail(f"{label}: no se muestran los 15 productos en el instrumento 1.")

    for row in rows(document.tables[4]):
        product, cd, si, en, isi = row[1], number(row[3]), number(row[4]), number(row[5]), number(row[6])
        if product not in balances or (si, en, cd) != (balances[product][0], balances[product][1], balances[product][3]):
            fail(f"{label}: ISI no usa el mismo balance en {product}.")
        if round(cd / (si + en) * 100, 2) != round(isi, 2) or not 0 <= isi <= 100:
            fail(f"{label}: ISI incorrecto en {product}.")

    for row in rows(document.tables[6]):
        product, dqs, dd, tqs = row[1], number(row[3]), number(row[4]), number(row[5])
        expected = expected_by_product.get(product)
        if expected is None or dd != 7 or not 0 <= dqs <= dd or round(dqs / dd * 100, 2) != round(tqs, 2):
            fail(f"{label}: TQS incorrecta en {product}.")

    for table_index in (8, 10, 14, 16):
        for row in rows(document.tables[table_index]):
            if row[1] not in expected_categories:
                fail(f"{label}: categoría inesperada {row[1]}.")
    for row in rows(document.tables[8]):
        category, vd, vpd, index = row[1], number(row[2]), number(row[3]), number(row[4])
        expected = expected_categories[category]
        if (round(vd, 2), round(vpd, 2), round(index, 2)) != (round(expected["actual_last"], 2), round(expected["actual_average"], 2), round(expected["daily_index"], 2)):
            fail(f"{label}: IE diario incorrecto en {category}.")
    for table_index in (10, 16):
        for row in rows(document.tables[table_index]):
            category, actual, forecast, mape = row[1], number(row[2]), number(row[3]), number(row[4])
            expected = expected_categories[category]
            if (round(actual, 2), round(forecast, 2), round(mape, 2)) != (round(expected["actual_total"], 2), round(expected["forecast_total"], 2), round(expected["mape"], 2)):
                fail(f"{label}: MAPE incorrecto en {category}.")
    for row in rows(document.tables[14]):
        category, vp, vpr, index = row[1], number(row[2]), number(row[3]), number(row[4])
        expected = expected_categories[category]
        if (round(vp, 2), round(vpr, 2), round(index, 2)) != (round(expected["actual_average"], 2), round(expected["reference_average"], 2), round(expected["pattern_index"], 2)):
            fail(f"{label}: IE de patrón incorrecto en {category}.")
    return expected_categories, expected_products


def main() -> None:
    pre_categories, pre_products = audit_document("pretest")
    post_categories, post_products = audit_document("posttest")

    pre_window = {(data["first_date"], data["last_date"]) for data in pre_categories.values()}
    post_window = {(data["first_date"], data["last_date"]) for data in post_categories.values()}
    if pre_window != post_window or len(pre_window) != 1:
        fail("Los escenarios no tienen exactamente el mismo período de siete días.")
    if [(item["product"], item["category"]) for item in pre_products] != [(item["product"], item["category"]) for item in post_products]:
        fail("Los escenarios no usan los mismos 15 productos.")
    shared_fields = ("actual", "cd", "si", "en", "sf", "dqs", "dd", "tqs")
    for pre, post in zip(pre_products, post_products):
        if any(pre[field] != post[field] for field in shared_fields):
            fail(f"La demanda o inventario de referencia no es común en {pre['product']}.")
    for category in pre_categories:
        if pre_categories[category]["actual_total"] != post_categories[category]["actual_total"]:
            fail(f"La demanda de referencia no es igual en {category}.")
        if not post_categories[category]["mape"] < pre_categories[category]["mape"]:
            fail(f"El escenario postest no mejora el MAPE simulado en {category}.")
    print("AUDITORÍA COMPARATIVA APROBADA")
    print("OK: ambos archivos tienen 8 instrumentos, 17 tablas y 15 productos.")
    print("OK: mismo período 16/09/2026–22/09/2026, mismas 105 observaciones y misma demanda de referencia simulada.")
    print("OK: CD, ISI, TQS, estacionalidad y MAPE cuadran; los saldos no son negativos.")
    print("AVISO: es una simulación controlada; no acredita una mejora real de la tienda.")


if __name__ == "__main__":
    main()
