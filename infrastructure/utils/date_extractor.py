import re
from datetime import datetime

def extract_date_from_text(text: str) -> datetime:
    """Busca 'FECHAI: dd/mm/yyyy' en el texto y devuelve un objeto datetime."""
    match = re.search(r'FECHAI\s*:?\s*(\d{1,2}/\d{1,2}/\d{4})', text, re.IGNORECASE)
    if not match:
        raise ValueError("No se encontró FECHAI en el encabezado del archivo")
    date_str = match.group(1)  # Ej: '23/04/2026'
    return datetime.strptime(date_str, '%d/%m/%Y')
