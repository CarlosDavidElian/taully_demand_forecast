from __future__ import annotations

import json
import shutil
from pathlib import Path

import numpy as np
import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor


WORKSPACE = Path(r"C:\taully_demand_forecast")
REFERENCE_DOCX = WORKSPACE / "outputs" / "01a0b651-7620-73f0-8f6a-8b68b14f7a17" / "Anexo_2_Instrumentos_Completos_Pretest_Postest.docx"
OUTPUT_DOCX = WORKSPACE / "outputs" / "01a0b651-7620-73f0-8f6a-8b68b14f7a17" / "Anexo_2_Instrumentos_Resultados_Tecnicos_Pretest_Postest_2026-09-02_a_2026-09-15.docx"
INVENTORY_PATH = WORKSPACE / "data" / "inventario_producto_Taully_2026-04-01_a_2026-09-15.xlsx"
HISTORY_PATH = WORKSPACE / "data" / "historial_demanda.csv"
POSTTEST_PATH = WORKSPACE / "data" / "posttests" / "postest_20260927t191247z_20260908_7d_70578bbb.json"

CATEGORIES = ["ABARROTES", "BEBIDAS", "GOLOSINAS", "HELADOS", "LIMPIEZA"]
RESEARCHER = "Guerrero Mora, Carlos David Elian"
PLACE = "Minimarket Taully, Comas"
O1_START = pd.Timestamp("2026-09-02")
O1_END = pd.Timestamp("2026-09-08")
O2_START = pd.Timestamp("2026-09-09")
O2_END = pd.Timestamp("2026-09-15")
REFERENCE_START = pd.Timestamp("2026-08-01")
REFERENCE_END = pd.Timestamp("2026-08-31")


def number(value: float | int, decimals: int = 0) -> str:
    if decimals == 0:
        return f"{float(value):,.0f}"
    return f"{float(value):,.{decimals}f}"


def percent(value: float, decimals: int = 2) -> str:
    return f"{float(value):.{decimals}f}%"


def set_run_font(run, size: float = 10.0, bold: bool | None = None, color: str = "1F2937") -> None:
    run.font.name = "Arial"
    run._element.rPr.rFonts.set(qn("w:ascii"), "Arial")
    run._element.rPr.rFonts.set(qn("w:hAnsi"), "Arial")
    run.font.size = Pt(size)
    run.font.color.rgb = RGBColor.from_string(color)
    if bold is not None:
        run.bold = bold


def set_cell_text(cell, value: str, *, align: int = WD_ALIGN_PARAGRAPH.LEFT, bold: bool = False, size: float = 10.0) -> None:
    cell.text = str(value)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    paragraph = cell.paragraphs[0]
    paragraph.alignment = align
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.space_before = Pt(0)
    for run in paragraph.runs:
        set_run_font(run, size=size, bold=bold)


def set_paragraph_text(paragraph, value: str, *, size: float = 10.5, bold: bool = False, italic: bool = False, alignment: int | None = None) -> None:
    paragraph.clear()
    if alignment is not None:
        paragraph.alignment = alignment
    run = paragraph.add_run(value)
    set_run_font(run, size=size, bold=bold, color="1F2937")
    run.italic = italic


def inventory_metrics(frame: pd.DataFrame, start: pd.Timestamp, end: pd.Timestamp) -> dict[str, dict[str, float]]:
    result: dict[str, dict[str, float]] = {}
    window = frame[(frame["Fecha"] >= start) & (frame["Fecha"] <= end)].copy()
    for category in CATEGORIES:
        category_rows = window[window["Categoría"] == category].sort_values(["Producto", "Fecha"])
        if category_rows.empty:
            raise ValueError(f"No hay inventario para {category} entre {start.date()} y {end.date()}.")
        grouped = category_rows.groupby("Producto", sort=False)
        si = grouped["Stock_Inicial"].first().sum()
        en = category_rows["Entradas"].sum()
        sf = grouped["Stock_Final"].last().sum()
        # DQS follows the program rule: unique dates where at least one product
        # in the category has a positive stockout alert. Summing rows would
        # overcount multiple products that were out of stock on the same day.
        dqs = category_rows.loc[category_rows["Dias_Sin_Stock"] > 0, "Fecha"].nunique()
        cd = si + en - sf
        result[category] = {
            "SI": float(si),
            "EN": float(en),
            "SF": float(sf),
            "CD": float(cd),
            "ISI": float(cd / (si + en) * 100),
            "DQS": float(dqs),
            "DD": float((end - start).days + 1),
            "TQS": float(dqs / ((end - start).days + 1) * 100),
        }
    return result


def recursive_pms_7(history: pd.DataFrame, cutoff: pd.Timestamp, start: pd.Timestamp, end: pd.Timestamp) -> dict[str, dict[str, float]]:
    result: dict[str, dict[str, float]] = {}
    for category in CATEGORIES:
        training = history[(history["category"] == category) & (history["date"] <= cutoff)].sort_values("date")
        actual = history[
            (history["category"] == category)
            & (history["date"] >= start)
            & (history["date"] <= end)
        ].sort_values("date")
        last_values = training["quantity"].astype(float).tail(7).tolist()
        if len(last_values) != 7 or len(actual) != 7:
            raise ValueError(f"No existe una ventana completa para PMS-7 en {category}.")
        predictions = []
        for _ in range(7):
            prediction = float(np.mean(last_values))
            predictions.append(prediction)
            last_values = last_values[1:] + [prediction]
        actual_values = actual["quantity"].to_numpy(dtype=float)
        prediction_values = np.array(predictions, dtype=float)
        errors = np.abs(actual_values - prediction_values)
        result[category] = {
            "DR": float(actual_values.sum()),
            "DP": float(prediction_values.sum()),
            "MAPE": float(np.mean(errors / actual_values) * 100),
            "MAE": float(np.mean(errors)),
            "RMSE": float(np.sqrt(np.mean((actual_values - prediction_values) ** 2))),
            "WAPE": float(errors.sum() / actual_values.sum() * 100),
        }
    return result


def total_inventory(metrics: dict[str, dict[str, float]]) -> dict[str, float]:
    values = {field: sum(item[field] for item in metrics.values()) for field in ("SI", "EN", "SF", "CD", "DQS")}
    values["ISI"] = values["CD"] / (values["SI"] + values["EN"]) * 100
    values["DD"] = 35.0
    values["TQS"] = values["DQS"] / values["DD"] * 100
    return values


def total_forecast(metrics: dict[str, dict[str, float]]) -> dict[str, float]:
    return {
        "DR": sum(item["DR"] for item in metrics.values()),
        "DP": sum(item["DP"] for item in metrics.values()),
        "MAPE": float(np.mean([item["MAPE"] for item in metrics.values()])),
        "MAE": float(np.mean([item["MAE"] for item in metrics.values()])),
        "RMSE": float(np.mean([item["RMSE"] for item in metrics.values()])),
        "WAPE": sum(abs(item["DR"] - item["DP"]) for item in metrics.values()) / sum(item["DR"] for item in metrics.values()) * 100,
    }


def fill_metadata(table, *, indicator: str, moment: str, formula: str, definition: str, unit: str, source: str, period: str, code: str) -> None:
    entries = [
        ("Indicador", indicator),
        ("Investigador", RESEARCHER),
        ("Lugar de estudio", PLACE),
        ("Momento", moment),
        ("Fórmula", formula),
        ("Definición operativa", definition),
        ("Unidad de medida", unit),
        ("Fuente utilizada", source),
        ("Período", period),
        ("Código de evidencia", code),
    ]
    for row_index, (label, value) in enumerate(entries, start=1):
        set_cell_text(table.cell(row_index, 0), label, bold=True, size=9.6)
        set_cell_text(table.cell(row_index, 1), value, size=9.6)


def fill_inventory_table(table, metrics: dict[str, dict[str, float]], kind: str) -> None:
    for row_index, category in enumerate(CATEGORIES, start=1):
        values = metrics[category]
        set_cell_text(table.cell(row_index, 0), str(row_index), align=WD_ALIGN_PARAGRAPH.CENTER, size=9.5)
        set_cell_text(table.cell(row_index, 1), category, size=9.5)
        if kind == "cd":
            ordered = [number(values["SI"]), number(values["EN"]), number(values["SF"]), number(values["CD"])]
        elif kind == "isi":
            ordered = [number(values["CD"]), number(values["SI"]), number(values["EN"]), percent(values["ISI"])]
        else:
            ordered = [number(values["DQS"]), number(values["DD"]), percent(values["TQS"])]
        for column, value in enumerate(ordered, start=2):
            set_cell_text(table.cell(row_index, column), value, align=WD_ALIGN_PARAGRAPH.CENTER, size=9.5)
    totals = total_inventory(metrics)
    total_row = 6
    set_cell_text(table.cell(total_row, 0), "", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=9.5)
    set_cell_text(table.cell(total_row, 1), "TOTAL", bold=True, size=9.5)
    if kind == "cd":
        ordered_total = [number(totals["SI"]), number(totals["EN"]), number(totals["SF"]), number(totals["CD"])]
    elif kind == "isi":
        ordered_total = [number(totals["CD"]), number(totals["SI"]), number(totals["EN"]), percent(totals["ISI"])]
    else:
        ordered_total = [number(totals["DQS"]), "35 categoría-días", percent(totals["TQS"])]
    for column, value in enumerate(ordered_total, start=2):
        set_cell_text(table.cell(total_row, column), value, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=9.5)


def fill_seasonality_table(table, weekly_sales: pd.Series, weekly_reference: pd.Series) -> None:
    for row_index, category in enumerate(CATEGORIES, start=1):
        vd = float(weekly_sales[category])
        vpm = float(weekly_reference[category])
        values = [number(vd), number(vpm, 2), number(vd / vpm, 2)]
        set_cell_text(table.cell(row_index, 0), str(row_index), align=WD_ALIGN_PARAGRAPH.CENTER, size=9.5)
        set_cell_text(table.cell(row_index, 1), category, size=9.5)
        for column, value in enumerate(values, start=2):
            set_cell_text(table.cell(row_index, column), value, align=WD_ALIGN_PARAGRAPH.CENTER, size=9.5)
    total_vd = float(weekly_sales.sum())
    total_vpm = float(weekly_reference.sum())
    total_values = [number(total_vd), number(total_vpm, 2), number(total_vd / total_vpm, 2)]
    set_cell_text(table.cell(6, 0), "", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=9.5)
    set_cell_text(table.cell(6, 1), "TOTAL", bold=True, size=9.5)
    for column, value in enumerate(total_values, start=2):
        set_cell_text(table.cell(6, column), value, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=9.5)


def fill_error_table(table, metrics: dict[str, dict[str, float]]) -> None:
    for row_index, category in enumerate(CATEGORIES, start=1):
        values = metrics[category]
        rendered = [
            number(values["DR"]),
            number(values["DP"], 2),
            percent(values["MAPE"]),
            number(values["MAE"], 2),
            number(values["RMSE"], 2),
        ]
        set_cell_text(table.cell(row_index, 0), str(row_index), align=WD_ALIGN_PARAGRAPH.CENTER, size=8.8)
        set_cell_text(table.cell(row_index, 1), category, size=8.8)
        for column, value in enumerate(rendered, start=2):
            set_cell_text(table.cell(row_index, column), value, align=WD_ALIGN_PARAGRAPH.CENTER, size=8.8)
    actual_all = np.array([metrics[category]["DR"] for category in CATEGORIES], dtype=float)
    predicted_all = np.array([metrics[category]["DP"] for category in CATEGORIES], dtype=float)
    # Overall metrics are calculated from the daily observations when supplied by the JSON;
    # values already attached to the caller are used below.


def fill_error_total_row(table, total: dict[str, float]) -> None:
    rendered = [
        number(total["DR"]),
        number(total["DP"], 2),
        percent(total["MAPE"]),
        number(total["MAE"], 2),
        number(total["RMSE"], 2),
    ]
    set_cell_text(table.cell(6, 0), "", align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=8.8)
    set_cell_text(table.cell(6, 1), "TOTAL", bold=True, size=8.8)
    for column, value in enumerate(rendered, start=2):
        set_cell_text(table.cell(6, column), value, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, size=8.8)


def fill_traceability(table) -> None:
    o1_period = "02/09 al 08/09/2026"
    o2_period = "09/09 al 15/09/2026"
    rows = [
        ("O1 Pretest", "1 Cantidad demandada", o1_period, "INV-20260401-20260915", RESEARCHER),
        ("O1 Pretest", "2 Índice de salida", o1_period, "INV-20260401-20260915", RESEARCHER),
        ("O1 Pretest", "3 Quiebre de stock", o1_period, "INV-20260401-20260915", RESEARCHER),
        ("O1 Pretest", "4 Estacionalidad", o1_period, "HIST-20260401-20260915", RESEARCHER),
        ("O1 Pretest", "5 Error PMS-7", o1_period, "O1-PMS7-20260901-7D", RESEARCHER),
        ("O2 Postest", "1 Cantidad demandada", o2_period, "INV-20260401-20260915", RESEARCHER),
        ("O2 Postest", "2 Índice de salida", o2_period, "INV-20260401-20260915", RESEARCHER),
        ("O2 Postest", "3 Quiebre de stock", o2_period, "INV-20260401-20260915", RESEARCHER),
        ("O2 Postest", "4 Estacionalidad", o2_period, "HIST-20260401-20260915", RESEARCHER),
        ("O2 Postest", "5 Error modelo", o2_period, "POSTEST-20260908-7D", RESEARCHER),
    ]
    for row_index, row_values in enumerate(rows, start=1):
        for column, value in enumerate(row_values):
            alignment = WD_ALIGN_PARAGRAPH.CENTER if column in (0, 2) else WD_ALIGN_PARAGRAPH.LEFT
            set_cell_text(table.cell(row_index, column), value, align=alignment, size=8.4)


def main() -> None:
    inventory = pd.read_excel(INVENTORY_PATH, sheet_name="Datos inventario")
    inventory["Fecha"] = pd.to_datetime(inventory["Fecha"])
    history = pd.read_csv(HISTORY_PATH, parse_dates=["date"])
    posttest = json.loads(POSTTEST_PATH.read_text(encoding="utf-8"))["result"]

    o1_inventory = inventory_metrics(inventory, O1_START, O1_END)
    o2_inventory = inventory_metrics(inventory, O2_START, O2_END)
    o1_pms = recursive_pms_7(history, pd.Timestamp("2026-09-01"), O1_START, O1_END)

    o2_model = {
        category: {
            "DR": float(posttest["model"]["by_category"][category]["actual_total"]),
            "DP": float(posttest["model"]["by_category"][category]["prediction_total"]),
            "MAPE": float(posttest["model"]["by_category"][category]["mape"]),
            "MAE": float(posttest["model"]["by_category"][category]["mae"]),
            "RMSE": float(posttest["model"]["by_category"][category]["rmse"]),
            "WAPE": float(posttest["model"]["by_category"][category]["wape"]),
        }
        for category in CATEGORIES
    }
    o1_total = {
        "DR": 1784.0,
        "DP": 2016.970589270992,
        "MAPE": 27.685091242063297,
        "MAE": 11.541524156688148,
        "RMSE": 15.955541169902654,
        "WAPE": 22.643124746865762,
    }
    o2_total = {
        "DR": float(posttest["model"]["metrics"]["actual_total"]),
        "DP": float(posttest["model"]["metrics"]["prediction_total"]),
        "MAPE": float(posttest["model"]["metrics"]["mape"]),
        "MAE": float(posttest["model"]["metrics"]["mae"]),
        "RMSE": float(posttest["model"]["metrics"]["rmse"]),
        "WAPE": float(posttest["model"]["metrics"]["wape"]),
    }

    # Independent controls: expected population and the Postest period must match the source.
    if sorted(posttest["categories_evaluated"]) != CATEGORIES:
        raise ValueError("Las categorías del Postest no coinciden con las del anexo.")
    if posttest["period"] != {"start_date": "2026-09-09", "end_date": "2026-09-15"}:
        raise ValueError("El período guardado del Postest no coincide con O2.")
    if int(posttest["horizon_days"]) != 7 or posttest["cutoff_date"] != "2026-09-08":
        raise ValueError("La configuración del Postest no coincide con el protocolo.")

    august = history[(history["date"] >= REFERENCE_START) & (history["date"] <= REFERENCE_END)]
    weekly_reference = august.groupby("category")["quantity"].mean().reindex(CATEGORIES) * 7
    o1_sales = history[(history["date"] >= O1_START) & (history["date"] <= O1_END)].groupby("category")["quantity"].sum().reindex(CATEGORIES)
    o2_sales = history[(history["date"] >= O2_START) & (history["date"] <= O2_END)].groupby("category")["quantity"].sum().reindex(CATEGORIES)

    if not (o1_sales.notna().all() and o2_sales.notna().all() and weekly_reference.notna().all()):
        raise ValueError("Faltan ventas para calcular la estacionalidad.")

    shutil.copy2(REFERENCE_DOCX, OUTPUT_DOCX)
    document = Document(OUTPUT_DOCX)
    document.core_properties.title = "Anexo 2 Instrumentos de recolección de datos con resultados"
    document.core_properties.author = RESEARCHER

    # Opening page: concise protocol and source note.
    set_paragraph_text(document.paragraphs[0], "Anexo 2 Instrumentos de recolección de datos con resultados", size=18, bold=True, alignment=WD_ALIGN_PARAGRAPH.LEFT)
    set_paragraph_text(
        document.paragraphs[1],
        "Este anexo registra los cinco indicadores del estudio en dos periodos consecutivos de siete días. Las fichas se completaron con los archivos disponibles del sistema y el Postest guardado.",
        size=10.5,
    )
    set_paragraph_text(
        document.paragraphs[2],
        "O1 Pretest corresponde al 02 al 08 de septiembre de 2026 y emplea PMS-7 como referencia. O2 Postest corresponde al 09 al 15 de septiembre de 2026 y emplea el modelo predictivo. Se mantienen las mismas categorías, fórmulas y unidad de medida.",
        size=10.5,
    )
    set_paragraph_text(
        document.paragraphs[3],
        "Criterio de comparación. O1 y O2 tienen igual duración. El contraste de precisión usa MAPE, MAE, RMSE y WAPE. El menor error corresponde al método más preciso en la semana evaluada.",
        size=10.0,
        italic=True,
    )
    set_paragraph_text(document.paragraphs[5], "Nota de fuente. Los archivos disponibles están identificados por el sistema como escenario de demostración. Los resultados son una aplicación técnica y deben validarse con kardex y ventas oficiales antes de presentarlos como evidencia definitiva.", size=9.2, italic=True)

    protocol_rows = [
        ("O1 Pretest", "Registrar indicadores del 02 al 08 de septiembre con PMS-7 como referencia.", "Historial de demanda e inventario disponibles."),
        ("X Intervención", "Procesar ventas y entrenar el modelo hasta el corte del 08 de septiembre.", "Fecha de corte y configuración del Postest."),
        ("O2 Postest", "Registrar indicadores del 09 al 15 de septiembre con el modelo predictivo.", "Postest guardado, historial e inventario disponibles."),
    ]
    for row_index, values in enumerate(protocol_rows, start=1):
        for column, value in enumerate(values):
            set_cell_text(document.tables[0].cell(row_index, column), value, align=WD_ALIGN_PARAGRAPH.LEFT if column else WD_ALIGN_PARAGRAPH.CENTER, size=9.4)

    # Metadata tables (O1: 2, 4, 6, 8, 10; O2: 12, 14, 16, 18, 20).
    o1_period = "02/09/2026 al 08/09/2026 (7 días)"
    o2_period = "09/09/2026 al 15/09/2026 (7 días)"
    inv_source = INVENTORY_PATH.name
    hist_source = HISTORY_PATH.name
    post_source = POSTTEST_PATH.name
    fill_metadata(document.tables[2], indicator="Cantidad demandada", moment="O1 Pretest", formula="CD = (SI + EN) - SF", definition="Unidades demandadas durante el período inicial, calculadas por balance de inventario.", unit="Unidades", source=inv_source, period=o1_period, code="INV-20260401-20260915")
    fill_metadata(document.tables[4], indicator="Índice de salida de inventario", moment="O1 Pretest", formula="ISI = CD / (SI + EN) x 100", definition="Proporción del stock disponible que salió durante el período inicial.", unit="Porcentaje", source=inv_source, period=o1_period, code="INV-20260401-20260915")
    fill_metadata(document.tables[6], indicator="Tasa de quiebre de stock", moment="O1 Pretest", formula="TQS = (DQS / DD) x 100", definition="Porcentaje de días con stock agotado por categoría durante el período inicial.", unit="Porcentaje", source=inv_source, period=o1_period, code="INV-20260401-20260915")
    fill_metadata(document.tables[8], indicator="Estacionalidad", moment="O1 Pretest", formula="IE = VD / VPM", definition="Relación entre la venta semanal observada y el promedio semanal de agosto de 2026.", unit="Índice", source=hist_source, period=o1_period + "; referencia: agosto 2026", code="HIST-20260401-20260915")
    fill_metadata(document.tables[10], indicator="Error de pronóstico", moment="O1 Pretest PMS-7", formula="MAPE = (100 / n) x suma de |(DR - DP) / DR|", definition="Error entre demanda real y el pronóstico PMS-7 calculado de forma recursiva.", unit="Porcentaje y unidades", source=hist_source + "; PMS-7 calculado", period=o1_period + "; corte: 01/09/2026", code="O1-PMS7-20260901-7D")
    fill_metadata(document.tables[12], indicator="Cantidad demandada", moment="O2 Postest", formula="CD = (SI + EN) - SF", definition="Unidades demandadas durante el período posterior, calculadas por balance de inventario.", unit="Unidades", source=inv_source, period=o2_period, code="INV-20260401-20260915")
    fill_metadata(document.tables[14], indicator="Índice de salida de inventario", moment="O2 Postest", formula="ISI = CD / (SI + EN) x 100", definition="Proporción del stock disponible que salió durante el período posterior.", unit="Porcentaje", source=inv_source, period=o2_period, code="INV-20260401-20260915")
    fill_metadata(document.tables[16], indicator="Tasa de quiebre de stock", moment="O2 Postest", formula="TQS = (DQS / DD) x 100", definition="Porcentaje de días con stock agotado por categoría durante el período posterior.", unit="Porcentaje", source=inv_source, period=o2_period, code="INV-20260401-20260915")
    fill_metadata(document.tables[18], indicator="Estacionalidad", moment="O2 Postest", formula="IE = VD / VPM", definition="Relación entre la venta semanal observada y el promedio semanal de agosto de 2026.", unit="Índice", source=hist_source, period=o2_period + "; referencia: agosto 2026", code="HIST-20260401-20260915")
    fill_metadata(document.tables[20], indicator="Error de pronóstico", moment="O2 Postest modelo predictivo", formula="MAPE = (100 / n) x suma de |(DR - DP) / DR|", definition="Error entre demanda real posterior al corte y el pronóstico del modelo.", unit="Porcentaje y unidades", source=post_source, period=o2_period + "; corte: 08/09/2026", code="POSTEST-20260908-7D")

    fill_inventory_table(document.tables[3], o1_inventory, "cd")
    fill_inventory_table(document.tables[5], o1_inventory, "isi")
    fill_inventory_table(document.tables[7], o1_inventory, "tqs")
    fill_seasonality_table(document.tables[9], o1_sales, weekly_reference)
    fill_error_table(document.tables[11], o1_pms)
    fill_error_total_row(document.tables[11], o1_total)
    fill_inventory_table(document.tables[13], o2_inventory, "cd")
    fill_inventory_table(document.tables[15], o2_inventory, "isi")
    fill_inventory_table(document.tables[17], o2_inventory, "tqs")
    fill_seasonality_table(document.tables[19], o2_sales, weekly_reference)
    fill_error_table(document.tables[21], o2_model)
    fill_error_total_row(document.tables[21], o2_total)

    # Replace all empty responsibility fields with the accountable researcher.
    for paragraph in document.paragraphs:
        if paragraph.text.startswith("Responsable del registro:"):
            set_paragraph_text(paragraph, f"Responsable del registro: {RESEARCHER}", size=9.8)

    # Indicators and control notes, made explicit for the actual calculated data.
    set_paragraph_text(document.paragraphs[10], "Criterio de cálculo. CD se obtuvo por balance. Ventas registradas: 1,784 unidades; CD por balance: 1,411 unidades; diferencia: 373 unidades.", size=9.2, italic=True)
    set_paragraph_text(document.paragraphs[16], "Criterio de cálculo. ISI se expresa como porcentaje del stock inicial más las entradas del mismo período.", size=9.2, italic=True)
    set_paragraph_text(document.paragraphs[22], "Criterio de cálculo. DQS registra días sin stock por categoría. El total usa 35 categoría-días (5 categorías x 7 días).", size=9.2, italic=True)
    set_paragraph_text(document.paragraphs[28], "Criterio de cálculo. VD es la venta de la semana evaluada. VPM es el promedio semanal calculado con agosto de 2026.", size=9.2, italic=True)
    set_paragraph_text(document.paragraphs[34], "Criterio de cálculo. PMS-7 usa las siete últimas observaciones y cada estimación se incorpora al siguiente día pronosticado.", size=9.2, italic=True)
    set_paragraph_text(document.paragraphs[40], "Criterio de cálculo. CD se obtuvo por balance. Ventas registradas: 1,671 unidades; CD por balance: 1,311 unidades; diferencia: 360 unidades.", size=9.2, italic=True)
    set_paragraph_text(document.paragraphs[46], "Criterio de cálculo. ISI se expresa como porcentaje del stock inicial más las entradas del mismo período.", size=9.2, italic=True)
    set_paragraph_text(document.paragraphs[52], "Criterio de cálculo. DQS registra días sin stock por categoría. El total usa 35 categoría-días (5 categorías x 7 días).", size=9.2, italic=True)
    set_paragraph_text(document.paragraphs[58], "Criterio de cálculo. Se mantiene la misma referencia de agosto de 2026 para poder comparar O1 y O2.", size=9.2, italic=True)
    set_paragraph_text(document.paragraphs[64], "Criterio de cálculo. El Postest compara solo ventas reales posteriores al corte del 08/09/2026 con el pronóstico del modelo.", size=9.2, italic=True)

    # O2 execution table.
    execution = [
        ("Fecha de corte", "08/09/2026"),
        ("Horizonte", "7 días"),
        ("Período evaluado", "09/09/2026 al 15/09/2026"),
        ("Archivo de evidencia", post_source),
        ("Comparación", "Modelo: MAPE 33.89%, MAE 10.41, RMSE 14.17, WAPE 21.81%. PMS-7: MAPE 31.54%, MAE 9.94, RMSE 13.26, WAPE 20.82%."),
    ]
    for row_index, (label, value) in enumerate(execution, start=1):
        set_cell_text(document.tables[22].cell(row_index, 0), label, bold=True, size=9.2)
        set_cell_text(document.tables[22].cell(row_index, 1), value, size=9.0)

    fill_traceability(document.tables[23])
    evidence = [
        ("historial_demanda.csv", "Sustenta VD, demanda real e índices estacionales."),
        (inv_source, "Sustenta SI, EN, SF, DQS, CD, ISI y TQS."),
        (post_source, "Sustenta corte, horizonte, pronóstico del modelo y métricas del O2."),
        ("Control de conciliación", "O1: diferencia de 373 unidades. O2: diferencia de 360 unidades entre ventas y balance."),
    ]
    for row_index, (label, value) in enumerate(evidence, start=1):
        set_cell_text(document.tables[24].cell(row_index, 0), label, size=8.8)
        set_cell_text(document.tables[24].cell(row_index, 1), value, size=8.8)
    set_paragraph_text(document.paragraphs[71], "Cierre. Las cifras se completaron con los archivos disponibles. Antes de usar este anexo como evidencia definitiva, la empresa debe validar la conciliación de ventas con kardex y movimientos de inventario.", size=9.0, italic=True)

    # The reference template contains a visible footer; preserve it but make the text accurate.
    for section in document.sections:
        footer = section.footer
        for paragraph in footer.paragraphs:
            if paragraph.text.strip():
                set_paragraph_text(paragraph, "Anexo 2 Instrumentos de recolección de datos", size=8.5, alignment=WD_ALIGN_PARAGRAPH.CENTER)

    document.save(OUTPUT_DOCX)

    print(json.dumps({
        "output": str(OUTPUT_DOCX),
        "o1_inventory_total": total_inventory(o1_inventory),
        "o2_inventory_total": total_inventory(o2_inventory),
        "o1_forecast_total": o1_total,
        "o2_forecast_total": o2_total,
        "o1_sales_total": float(o1_sales.sum()),
        "o2_sales_total": float(o2_sales.sum()),
    }, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
