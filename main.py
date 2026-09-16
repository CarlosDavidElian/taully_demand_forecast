import sys
from pathlib import Path

from application.services.catalog_service import CatalogService
from application.services.category_history_rebuild_service import CategoryHistoryRebuildService
from interfaces.cli import parse_arguments
from infrastructure.repositories.csv_repository import CSVDemandRepository
from infrastructure.repositories.catalog_repository import ExcelCatalogRepository
from infrastructure.readers.excel_reader import ExcelReader
from application.services.ingest_service import IngestService
from application.services.train_service import TrainService
from application.services.predict_service import PredictService
from application.use_cases.process_report import ProcessReportUseCase
from application.use_cases.run_forecast import RunForecastUseCase
from config.settings import CATEGORY_PRODUCT_MIX_FILE, DATA_DIR, HISTORY_BACKUPS_DIR, UNCATEGORIZED_SALES_FILE

def main():
    args = parse_arguments()

    # Inicializar solo las dependencias requeridas por el comando. Así el
    # entrenamiento y el pronóstico no dependen de abrir el catálogo Excel.
    demand_repo = CSVDemandRepository()
    train_service = TrainService(demand_repo)
    predict_service = PredictService(demand_repo)

    if args.command == 'load':
        ingest_service = IngestService(demand_repo, CatalogService(ExcelCatalogRepository()))
        use_case = ProcessReportUseCase(ingest_service)
        try:
            demands = use_case.execute(args.file)
            print("Resumen de demanda procesada:")
            for d in demands[:5]:  # Mostrar solo los primeros 5
                print(f"  - {d.date.strftime('%Y-%m-%d')} | {d.category}: {d.quantity:.2f}")
            if len(demands) > 5:
                print(f"  ... y {len(demands)-5} registros más.")
            unmatched = ingest_service.last_catalog_summary["unmatched_products"]
            if unmatched:
                print("Productos sin coincidencia en el catálogo:")
                for product in unmatched:
                    print(f"  - {product}")
        except Exception as e:
            print(f"Error al procesar el archivo: {e}")
            sys.exit(1)

    elif args.command == 'train':
        try:
            metrics = train_service.run()
            print(
                f"Modelo entrenado para {train_service.last_training_summary['trained_categories']} categorías. "
                f"WAPE promedio: {metrics.get('wape', 'N/A')}%"
            )
        except Exception as e:
            print(f"Error al entrenar: {e}")
            sys.exit(1)

    elif args.command == 'rebuild-history':
        try:
            service = CategoryHistoryRebuildService(
                demand_repo=demand_repo,
                catalog_service=CatalogService(ExcelCatalogRepository()),
                sale_reader=ExcelReader(),
                backup_directory=HISTORY_BACKUPS_DIR,
                uncategorized_sales_path=UNCATEGORIZED_SALES_FILE,
                category_product_mix_path=CATEGORY_PRODUCT_MIX_FILE,
            )
            result = service.rebuild(DATA_DIR.glob('Reporte_Taully_*.xlsx'))
            print(
                f"Historial reconstruido con {result.reports} reportes, {result.dates} días, "
                f"{result.categories} categorías y {result.demand_records} observaciones."
            )
            print(
                f"Cobertura categorizada: {result.categorized_units:.0f} de {result.source_units:.0f} unidades "
                f"({result.categorized_unit_share:.1%})."
            )
            print(f"Productos sin categoría: {result.uncategorized_products}. Revisa {result.uncategorized_sales_path}.")
            print(f"Relación de productos por categoría: {result.category_product_mix_path}")
            if result.backup_path is not None:
                print(f"Respaldo del historial anterior: {result.backup_path}")
        except Exception as e:
            print(f"Error al reconstruir el historial: {e}")
            sys.exit(1)

    elif args.command == 'predict':
        use_case = RunForecastUseCase(predict_service)
        try:
            predictions = use_case.execute(args.days)
        except Exception as e:
            print(f"Error al predecir: {e}")
            sys.exit(1)

    else:
        print("Comando no reconocido. Use: load, rebuild-history, train o predict")
        sys.exit(1)

if __name__ == "__main__":
    main()
