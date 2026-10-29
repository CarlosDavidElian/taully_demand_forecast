"""Tablas de las fichas, conservando unidades y referencia explícitas."""
from collections import defaultdict
from datetime import date, timedelta


def build_indicator_tables(result, demands):
    inventory = result.get("inventory", {})
    metrics = inventory.get("metrics", []) if inventory.get("available") else []
    tables = []

    def add(title, formula, headers, rows):
        tables.append(dict(title=title, formula=formula, headers=headers, rows=rows))

    note = inventory.get("message", "Sin inventario aplicable al período.")
    if inventory.get("inventory", {}).get("is_demonstration"):
        note += " Fuente demostrativa; pendiente validar con kardex real."
    add("1. Cantidad demandada", "CD = SI + EN − SF. " + note,
        ["Categoría", "SI", "EN", "SF", "CD"],
        [[m["category"], m["SI"], m["EN"], m["SF"], m["CD"]] for m in metrics])
    add("2. Índice de salida de inventario", "ISI = CD / (SI + EN). Se presenta en porcentaje.",
        ["Categoría", "CD", "Stock disponible", "ISI (%)"],
        [[m["category"], m["CD"], m["SI"] + m["EN"], m["ISI"]] for m in metrics])
    add("3. Tasa de quiebre de stock", "TQS = DQS / DD × 100. Se cuentan fechas distintas por categoría.",
        ["Categoría", "DQS", "DD", "TQS (%)"],
        [[m["category"], m["DQS"], m["DD"], m["TQS"]] for m in metrics])

    # Referencia de las fichas del proyecto: junio de 2026, no se sustituye silenciosamente.
    start, end = date(2026, 6, 1), date(2026, 6, 30)
    reference, dates = defaultdict(float), defaultdict(set)
    for demand in demands:
        day = demand.date.date()
        if start <= day <= end:
            reference[demand.category] += float(demand.quantity)
            dates[demand.category].add(day)
    categories = result["categories_evaluated"]
    valid = end <= date.fromisoformat(result["cutoff_date"]) and all(len(dates[c]) == 30 for c in categories)
    horizon = result["horizon_days"]
    note = "Referencia: junio de 2026. VPR = ventas de junio / 30 × días evaluados. IE = VP / VPR."
    if not valid:
        note += " Referencia incompleta o posterior al corte: no se calcula IE."
    rows = []
    for c in categories:
        vp = result["model"]["by_category"][c]["actual_total"]
        vpr = reference[c] / 30 * horizon if valid else None
        rows.append([c, vp, vpr, vp / vpr if vpr else None])
    vp_total = sum(r[1] for r in rows)
    vpr_total = sum(reference[c] for c in categories) / 30 * horizon if valid else None
    rows.append(["GLOBAL", vp_total, vpr_total, vp_total / vpr_total if vpr_total else None])
    add("4. Estacionalidad por categoría", note, ["Categoría", "VP", "VPR", "IE"], rows)
    daily = defaultdict(float)
    for row in result["observations"]:
        daily[row["date"]] += row["actual_demand"]
    daily_reference = sum(reference[c] for c in categories) / 30 if valid else None
    add("4b. Estacionalidad diaria", "IE diario = venta diaria / promedio diario de junio de 2026. No demuestra por sí solo recurrencia estacional.",
        ["Fecha", "Venta diaria", "Referencia diaria", "IE diario"],
        [[day, value, daily_reference, value / daily_reference if daily_reference else None] for day, value in sorted(daily.items())])
    rows = []
    for c in categories + ["GLOBAL"]:
        model = result["model"]["metrics"] if c == "GLOBAL" else result["model"]["by_category"][c]
        base = result["baseline"]["metrics"] if c == "GLOBAL" else result["baseline"]["by_category"][c]
        rows.append([c, model.get("actual_total"), base.get("prediction_total"), base.get("mape"), model.get("prediction_total"), model.get("mape")])
    add("5. Error de pronóstico", "MAPE = promedio de |DR − DP| / DR × 100 por observación diaria. Para las fichas: pretest usa PMS-7; postest usa modelo. GLOBAL no es una categoría adicional.",
        ["Categoría", "DR", "DP PMS-7", "MAPE PMS-7 (%)", "DP modelo", "MAPE modelo (%)"], rows)
    return tables
