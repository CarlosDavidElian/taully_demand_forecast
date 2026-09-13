from application.services.predict_service import PredictService

class RunForecastUseCase:
    """Genera pronósticos con un modelo previamente entrenado."""

    def __init__(self, predict_service: PredictService):
        self.predict_service = predict_service

    def execute(self, days: int = 7):
        print(f"Generando predicciones para {days} días...")
        predictions = self.predict_service.predict_future(days)
        
        print("\nPredicciones por categoría:")
        for cat, demands in predictions.items():
            print(f"\n  {cat}:")
            for d in demands:
                print(f"    {d.date.strftime('%Y-%m-%d')}: {d.quantity:.2f} unidades")
        
        return predictions
