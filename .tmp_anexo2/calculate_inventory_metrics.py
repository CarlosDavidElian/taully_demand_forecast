"""Resume los cálculos de inventario disponibles por categoría.

No modifica los archivos de entrada. Este resumen deja explícita la base
usada para las tres fichas de inventario del Anexo 2.
"""

from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "data" / "inventario_producto_Taully_2026-04-01_a_2026-09-15.xlsx"


def main() -> None:
    data = pd.read_excel(SOURCE)
    data["Fecha"] = pd.to_datetime(data["Fecha"])
    total_days = data["Fecha"].nunique()
    print(
        f"filas={len(data)}; fechas={total_days}; "
        f"rango={data['Fecha'].min():%d/%m/%Y}–{data['Fecha'].max():%d/%m/%Y}"
    )
    print("categoria|SI|EN|SF|ventas|CD_formula|ISI_pct|DQS|DD|TQS_pct")
    for category, group in data.groupby("Categoría", sort=True):
        stock_initial = int(group["Stock_Inicial"].sum())
        entries = int(group["Entradas"].sum())
        stock_final = int(group["Stock_Final"].sum())
        sales = int(group["Cantidad_Vendida"].sum())
        demand_formula = stock_initial + entries - stock_final
        isi = 100 * demand_formula / (stock_initial + entries)
        stockout_days = int(group.loc[group["Dias_Sin_Stock"] > 0, "Fecha"].nunique())
        tqs = 100 * stockout_days / total_days
        print(
            f"{category}|{stock_initial}|{entries}|{stock_final}|{sales}|"
            f"{demand_formula}|{isi:.6f}|{stockout_days}|{total_days}|{tqs:.6f}"
        )
    print(
        "total|{}|{}|{}|{}|{}".format(
            int(data["Stock_Inicial"].sum()),
            int(data["Entradas"].sum()),
            int(data["Stock_Final"].sum()),
            int(data["Cantidad_Vendida"].sum()),
            int((data["Stock_Inicial"] + data["Entradas"] - data["Stock_Final"]).sum()),
        )
    )


if __name__ == "__main__":
    main()
