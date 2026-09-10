import sys
from pathlib import Path

from interfaces.cli import parse_arguments
from infrastructure.repositories.csv_repository import CSVDemandRepository
from infrastructure.repositories.catalog_repository import ExcelCatalogRepository
from application.services.ingest_service import IngestService
from application.services.train_service import TrainService
from application.services.predict_service import PredictService
from application.use_cases.process_report import ProcessReportUseCase
from application.use_cases.run_forecast import RunForecastUseCase

def main():
    args = parse_arguments()

    # Inicializar solo las dependencias requeridas por el comando. Así el
    # entrenamiento y el pronóstico no dependen de abrir el catálogo Excel.
    demand_repo = CSVDemandRepository()
    train_service = TrainService(demand_repo)
    predict_service = PredictService(demand_repo)

    if args.command == 'load':
        catalog_repo = ExcelCatalogRepository()
        ingest_service = IngestService(catalog_repo, demand_repo)
        use_case = ProcessReportUseCase(ingest_service)
        try:
            demands = use_case.execute(args.file)
            print("Resumen de demanda procesada:")
            for d in demands[:5]:  # Mostrar solo los primeros 5
                print(f"  - {d.date.strftime('%Y-%m-%d')} | {d.category}: {d.quantity:.2f}")
            if len(demands) > 5:
                print(f"  ... y {len(demands)-5} registros más.")
        except Exception as e:
            print(f"Error al procesar el archivo: {e}")
            sys.exit(1)

    elif args.command == 'train':
        use_case = RunForecastUseCase(train_service, predict_service)
        try:
            # Solo entrenar, sin predecir (o predecir solo para validar)
            metrics = train_service.run()
            print(f"Modelo entrenado. MAPE: {metrics.get('mape', 'N/A')}%")
        except Exception as e:
            print(f"Error al entrenar: {e}")
            sys.exit(1)

    elif args.command == 'predict':
        use_case = RunForecastUseCase(train_service, predict_service)
        try:
            predictions = use_case.execute(args.days)
        except Exception as e:
            print(f"Error al predecir: {e}")
            sys.exit(1)

    else:
        print("Comando no reconocido. Use: load, train o predict")
        sys.exit(1)

if __name__ == "__main__":
    main()
