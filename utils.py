"""Utilidades compartidas por los módulos de transformación SCC y OCN."""

from datetime import datetime


def strip(value):
    if value is None:
        return ""
    return str(value).strip()


def format_fecha(value):
    """Devuelve la fecha en formato AAAAMMDD como texto."""
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y%m%d")
    s = str(value).strip()
    if len(s) == 8 and s.isdigit():
        return s
    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y%m%d")
        except ValueError:
            continue
    return s


def to_number(value):
    """Convierte a int si es un entero exacto, si no a float."""
    if value is None:
        return value
    s = str(value).strip()
    if s == "":
        return value
    try:
        f = float(s)
    except ValueError:
        return value
    if f.is_integer():
        return int(f)
    return f


def normalizar_item(value):
    """Quita ceros a la izquierda de códigos de ítem puramente numéricos.
    Algunos reportes traen el código con relleno de ceros (ej. '0005067')
    mientras que el ERP y la lista de precios lo manejan sin relleno
    ('5067'). Si el código trae letras, se deja tal cual."""
    s = strip(value)
    if s.isdigit():
        stripped = s.lstrip("0")
        return stripped if stripped else "0"
    return s


def nit_base(value):
    """Quita el dígito de verificación (después del guion) y espacios."""
    s = strip(value)
    return s.split("-")[0].strip()
