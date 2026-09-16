import argparse
from pathlib import Path

def parse_arguments():
    parser = argparse.ArgumentParser(description="Sistema de Predicción de Demanda - Minimarket Taully")
    subparsers = parser.add_subparsers(dest='command', help='Comandos disponibles')

    # Comando: load
    parser_load = subparsers.add_parser('load', help='Cargar un reporte de ventas (Excel/PDF)')
    parser_load.add_argument('--file', required=True, help='Ruta del archivo (Excel o PDF)')

    # Comando: train
    parser_train = subparsers.add_parser('train', help='Entrenar el modelo con el historial actual')

    subparsers.add_parser(
        'rebuild-history',
        help='Reconstruir el historial diario por categoría desde los reportes POS',
    )

    # Comando: predict
    parser_predict = subparsers.add_parser('predict', help='Predecir demanda futura')
    parser_predict.add_argument('--days', type=int, default=7, help='Número de días a predecir (default: 7)')

    return parser.parse_args()
