"""Normalization helpers mirroring legacy VBA utility functions.

The functions in this module replicate the behaviour of the VBA helpers
found in ``Módulo4.bas``. They are intentionally permissive and try to
accept the same loose input formats that were tolerated by the macros:

* ``normalize`` removes accents and whitespace/punctuation to build a
  canonical comparison key for headers.
* ``date_only_ex2`` and ``time_to_sec_ex`` parse the flexible date/time
  formats that appeared in the spreadsheets (DMY/MDY, ``hh:mm`` or
  ``hhmm`` numbers, and Excel serial values).
* ``find_header_row`` and ``find_col_any_in_row`` reproduce the header
  detection heuristics used in Excel, while ``find_header_row_visitas``
  and ``find_header_row_cargas`` look for the specialised visit/load
  reports.
* ``normalize_vehiculo_key`` and ``servicio_id_from_components`` keep the
  same "vehicle-date-sequence" identifiers that the VBA produced.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
import os
import re
import unicodedata
from typing import Iterable, List, Optional, Sequence


ACCENT_REPLACEMENTS = str.maketrans(
    {
        "á": "a",
        "à": "a",
        "ä": "a",
        "â": "a",
        "Á": "a",
        "À": "a",
        "Ä": "a",
        "Â": "a",
        "é": "e",
        "è": "e",
        "ë": "e",
        "ê": "e",
        "É": "e",
        "È": "e",
        "Ë": "e",
        "Ê": "e",
        "í": "i",
        "ì": "i",
        "ï": "i",
        "î": "i",
        "Í": "i",
        "Ì": "i",
        "Ï": "i",
        "Î": "i",
        "ó": "o",
        "ò": "o",
        "ö": "o",
        "ô": "o",
        "Ó": "o",
        "Ò": "o",
        "Ö": "o",
        "Ô": "o",
        "ú": "u",
        "ù": "u",
        "ü": "u",
        "û": "u",
        "Ú": "u",
        "Ù": "u",
        "Ü": "u",
        "Û": "u",
        "ñ": "n",
        "Ñ": "n",
    }
)


def _strip_accents(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value.translate(ACCENT_REPLACEMENTS))
    return "".join(ch for ch in normalized if not unicodedata.combining(ch))


def normalize(value: str) -> str:
    """Replicate the VBA ``Normalize`` helper.

    The function removes accents, whitespace, and punctuation so that
    header comparisons are stable even when spreadsheets use slightly
    different labels.
    """

    if value is None:
        return ""
    text = str(value).strip().lower()
    text = _strip_accents(text)
    for token in (" ", "_", ".", "-", "/", "\\"):
        text = text.replace(token, "")
    return text


def find_header_row(
    data: "pandas.DataFrame",
    required_normalized_headers: Sequence[str] | None = None,
    *,
    max_rows: int = 25,
    default_row: int | None = 5,
) -> int:
    """Locate the header row similarly to ``FindHeaderRow`` in VBA.

    Args:
        data: DataFrame loaded with ``header=None``.
        required_normalized_headers: Expected normalized labels. When all
            are found in the same row the search stops.
        max_rows: Number of top rows to scan.
        default_row: Row index (0-based) to return when nothing matches.

    Returns:
        The 0-based header row index.
    """

    import pandas as pd  # imported lazily to keep light dependencies for callers

    if required_normalized_headers is None:
        required_normalized_headers = (
            normalize("Kilómetros"),
            normalize("Fecha Inicio"),
            normalize("Hora Inicio"),
        )

    max_rows = min(max_rows, len(data))
    for idx in range(max_rows):
        row = data.iloc[idx].fillna("")
        normalized = {normalize(v) for v in row.tolist()}
        if all(key in normalized for key in required_normalized_headers):
            return idx
    return default_row if default_row is not None else 0


def find_col_any_in_row(data: "pandas.DataFrame", header_row: int, aliases: Sequence[str]) -> int:
    """Return the first column index matching any of the aliases.

    The search is based on ``normalize`` and mirrors the permissive check
    implemented in ``FindColAnyInRow``.
    """

    if header_row < 0 or header_row >= len(data):
        return -1
    row = data.iloc[header_row].fillna("")
    normalized_aliases = {normalize(a) for a in aliases}
    for idx, value in enumerate(row.tolist()):
        if normalize(value) in normalized_aliases:
            return idx
    return -1


@dataclass(frozen=True)
class CatCategory:
    nombre: str

    @classmethod
    def from_raw(cls, raw: str) -> "CatCategory":
        normalized = normalize(raw)
        if "patio" in normalized:
            return cls("Patio")
        if "gaso" in normalized or "diesel" in normalized:
            return cls("Diesel")
        if "taller" in normalized or "mecan" in normalized or "servicio" in normalized:
            return cls("Taller")
        return cls("Otros")

    def guess_from_site(self, site: str) -> "CatCategory":
        guessed = CatCategory.from_raw(site)
        if self.nombre == "Otros" and guessed.nombre != "Otros":
            return guessed
        return self


EXCEL_EPOCH = datetime(1899, 12, 30)


def _excel_serial_to_date(value: float) -> date:
    return (EXCEL_EPOCH + timedelta(days=float(value))).date()


def date_only_ex2(value, order: str) -> Optional[date]:
    """Parse dates written as MDY/DMY strings or Excel serials.

    The behaviour mirrors ``DateOnlyEx2``: on invalid input ``None`` is
    returned instead of raising.
    """

    if value is None:
        return None
    if isinstance(value, date) and not isinstance(value, datetime):
        return value
    if isinstance(value, datetime):
        return value.date()

    try:
        numeric_value = float(value)
    except (TypeError, ValueError):
        numeric_value = None
    else:
        # Excel stores dates as a fraction of a day from the epoch.
        if numeric_value >= 1:
            return _excel_serial_to_date(numeric_value)

    text = str(value).strip()
    if not text:
        return None

    separators = ["-", "/"]
    for sep in separators:
        if sep in text:
            parts = text.split(sep)
            break
    else:
        return None

    if len(parts) != 3:
        return None

    try:
        order = order.strip().upper()
        if order == "MDY":
            m, d, y = (int(p) for p in parts)
        elif order == "DMY":
            d, m, y = (int(p) for p in parts)
        else:
            return None
    except ValueError:
        return None

    if y < 100:
        y += 2000

    try:
        return date(y, m, d)
    except ValueError:
        return None


def time_to_sec_ex(value) -> int:
    """Parse times written as ``hh:mm[:ss]`` or ``hhmm`` numbers.

    Invalid values return ``0`` as in the VBA version.
    """

    if value is None:
        return 0

    if isinstance(value, (int, float)):
        if float(value) < 1:
            return int(round(float(value) * 86400))
        # fall through to treat large numbers as HHMM

    text = str(value).strip()
    if not text:
        return 0

    if ":" in text:
        parts = text.split(":")
        try:
            h = int(parts[0])
            m = int(parts[1]) if len(parts) >= 2 else 0
            s = int(parts[2]) if len(parts) >= 3 else 0
        except ValueError:
            return 0
    else:
        try:
            number = int(text)
        except ValueError:
            return 0
        h, m = divmod(number, 100)
        s = 0

    if not (0 <= h <= 47 and 0 <= m <= 59 and 0 <= s <= 59):
        return 0

    return h * 3600 + m * 60 + s


def _seconds_to_time(seconds: int) -> time:
    seconds = max(0, int(seconds)) % (24 * 3600)
    h, remainder = divmod(seconds, 3600)
    m, s = divmod(remainder, 60)
    return time(hour=h, minute=m, second=s)


def normalize_vehiculo_key(value) -> str:
    """Return an upper-case alphanumeric vehicle key.

    Mirrors ``NormalizeVehiculoKey`` by removing any non-alphanumeric
    character. If the value cannot be converted to string an empty key is
    returned.
    """

    if value is None:
        return ""

    try:
        text = str(value).strip()
    except Exception:
        return ""

    if text.lower() == "none":
        return ""

    filtered = re.sub(r"[^0-9A-Za-z]", "", text)
    return filtered.upper()


def servicio_id_from_components(vehiculo: str, fecha, secuencial: int) -> str:
    veh_key = normalize_vehiculo_key(vehiculo) or "NA"
    if isinstance(fecha, (datetime, date)):
        fecha_key = fecha.strftime("%Y%m%d")
    else:
        try:
            fecha_dt = _excel_serial_to_date(float(fecha))
            fecha_key = fecha_dt.strftime("%Y%m%d")
        except Exception:
            fecha_key = "00000000"

    secuencial = max(0, int(secuencial))
    return f"{veh_key}-{fecha_key}-{secuencial:03d}"


def find_header_row_visitas(data: "pandas.DataFrame", default_row: int = 7) -> int:
    aliases = {
        "unidad",
        "economico",
        "economico#",
        "noeconomico",
        "unidadvehiculo",
    }
    fecha_aliases = {"fechallegada", "fecha llegada", "fechaarribo", "fechaarrive", "fllegada", "fecha"}
    hora_aliases = {"horallegada", "hora llegada", "hllegada", "horafecha", "hora"}
    cat_aliases = {"categoria", "categora", "categoriavisita", "tipovisita", "tipo", "categora visita"}

    for idx in range(min(100, len(data))):
        row = {normalize(v) for v in data.iloc[idx].fillna("").tolist()}
        if row & aliases and row & fecha_aliases and row & hora_aliases and row & cat_aliases:
            return idx
    return default_row


def find_header_row_cargas(data: "pandas.DataFrame", default_row: int = 0) -> int:
    unidad_aliases = {"unidad", "vehiculo", "vehculo", "carro"}
    fecha_aliases = {"fecha", "fregistro", "f registro", "fservicio"}
    division_aliases = {"division", "divisin", "div"}

    for idx in range(min(100, len(data))):
        row = {normalize(v) for v in data.iloc[idx].fillna("").tolist()}
        if row & unidad_aliases and row & fecha_aliases and row & division_aliases:
            return idx
    return default_row


def extraer_division_desde_nombre(file_name: str) -> str:
    base = os.path.splitext(os.path.basename(str(file_name).strip()))[0]
    if not base:
        return ""
    lower_base = base.lower()

    match = re.search(r"divisi[oó]n?\s*([\w]+)", lower_base)
    if match:
        identifier = match.group(1)
    else:
        match = re.search(r"^div\s*([\w]+)", lower_base)
        identifier = match.group(1) if match else base

    identifier = identifier.strip()
    if not identifier:
        return base
    if len(identifier) == 1:
        formatted = identifier.upper()
    else:
        formatted = identifier.capitalize()
    return f"Division {formatted}" if not formatted.lower().startswith("division") else formatted


# === DataFrame helpers ===

def _require_pandas():
    try:
        import pandas as pd  # type: ignore
    except ModuleNotFoundError as exc:
        raise ModuleNotFoundError(
            "pandas is required for Excel parsing functions; install it with 'pip install pandas openpyxl'."
        ) from exc
    return pd


def read_division_excel(path: str, *, date_order: str = "MDY", sheet_name: str | None = None):
    pd = _require_pandas()
    parsed_sheet = 0 if sheet_name is None else sheet_name
    raw = pd.read_excel(path, sheet_name=parsed_sheet, header=None)
    header_row = find_header_row(raw)

    col_km = find_col_any_in_row(raw, header_row, ["Kilómetros", "KMS", "Kilometros"])
    col_fecha = find_col_any_in_row(raw, header_row, ["Fecha Inicio", "F Servicio", "Fecha", "F_Servicio"])
    col_hora_ini = find_col_any_in_row(raw, header_row, ["Hora Inicio", "HoraInicial", "Inicio", "HI"])
    col_hora_fin = find_col_any_in_row(raw, header_row, ["Hora Fin", "HoraFinal", "Fin", "HF"])
    col_veh = find_col_any_in_row(raw, header_row, ["Vehiculo", "Vehículo", "Unidad", "Carro", "KCarro"])
    col_cliente = find_col_any_in_row(raw, header_row, ["Cliente / SiteVisit", "Cliente", "Cliente SiteVisit"])

    required = [col_km, col_fecha, col_hora_ini]
    if any(idx < 0 for idx in required):
        raise ValueError("No se pudieron detectar las columnas obligatorias (kilómetros, fecha, hora inicio)")

    records: List[dict] = []
    for _, row in raw.iloc[header_row + 1 :].iterrows():
        fecha = date_only_ex2(row[col_fecha], date_order)
        if not fecha:
            continue
        hora_ini_sec = time_to_sec_ex(row[col_hora_ini])
        hora_fin_sec = time_to_sec_ex(row[col_hora_fin]) if col_hora_fin >= 0 else hora_ini_sec
        vehiculo = normalize_vehiculo_key(row[col_veh]) if col_veh >= 0 else ""
        cliente = str(row[col_cliente]).strip() if col_cliente >= 0 and not pd.isna(row[col_cliente]) else ""
        km = float(row[col_km]) if pd.notna(row[col_km]) else 0.0

        start_dt = datetime.combine(fecha, _seconds_to_time(hora_ini_sec))
        end_dt = datetime.combine(fecha, _seconds_to_time(hora_fin_sec))
        minutes = max(0.0, (end_dt - start_dt).total_seconds() / 60.0)

        records.append(
            {
                "division": extraer_division_desde_nombre(path),
                "vehiculo": vehiculo,
                "inicio": start_dt,
                "fin": end_dt,
                "km": km,
                "minutos": minutes,
                "cliente_site": cliente,
                "servicio_id": servicio_id_from_components(vehiculo, fecha, len(records)),
                "tipo": "SERVICIO",
            }
        )

    return pd.DataFrame.from_records(records)


def read_visitas_excel(path: str, *, sheet_name: str | None = None, date_order: str = "DMY"):
    pd = _require_pandas()
    parsed_sheet = 0 if sheet_name is None else sheet_name
    raw = pd.read_excel(path, sheet_name=parsed_sheet, header=None)
    header_row = find_header_row_visitas(raw)

    col_unidad = find_col_any_in_row(raw, header_row, ["Unidad", "Económico", "Economico", "No Economico", "NoEconomico"])
    col_fecha = find_col_any_in_row(raw, header_row, ["Fecha Llegada", "FechaLlegada", "Fecha Arribo", "Fecha", "F Llegada"])
    col_hora = find_col_any_in_row(raw, header_row, ["Hora Llegada", "HoraLlegada", "Hora Arribo", "Hora"])
    col_fecha_sal = find_col_any_in_row(raw, header_row, ["Fecha Salida", "FechaSalida", "F Salida", "Fecha Fin"])
    col_hora_sal = find_col_any_in_row(raw, header_row, ["Hora Salida", "HoraSalida", "H Salida", "Hora Fin"])
    col_duracion = find_col_any_in_row(raw, header_row, ["Tiempo de Visita", "TiempoVisita", "Duración", "Duracion"])
    col_cat = find_col_any_in_row(raw, header_row, ["Categoría", "Categoria", "Categoría Visita", "Tipo", "Tipo Visita"])
    col_sitio = find_col_any_in_row(raw, header_row, ["Sitio", "Lugar", "Ubicacion", "Ubicación", "Punto", "Destino"])

    if col_unidad < 0 or col_fecha < 0 or col_hora < 0:
        raise ValueError("Visitas: faltan columnas mínimas (Unidad, Fecha/Hora Llegada)")

    records: List[dict] = []
    for _, row in raw.iloc[header_row + 1 :].iterrows():
        vehiculo_raw = row[col_unidad]
        vehiculo = normalize_vehiculo_key(vehiculo_raw)
        if not vehiculo:
            continue

        fecha = date_only_ex2(row[col_fecha], date_order)
        hora_sec = time_to_sec_ex(row[col_hora])
        if not fecha:
            continue

        fecha_sal = date_only_ex2(row[col_fecha_sal], date_order) if col_fecha_sal >= 0 else None
        hora_sal_sec = time_to_sec_ex(row[col_hora_sal]) if col_hora_sal >= 0 else 0
        duracion = time_to_sec_ex(row[col_duracion]) if col_duracion >= 0 else 0

        inicio_abs = datetime.combine(fecha, _seconds_to_time(hora_sec))
        if fecha_sal or hora_sal_sec:
            fecha_dest = fecha_sal or fecha
            fin_abs = datetime.combine(fecha_dest, _seconds_to_time(hora_sal_sec))
        elif duracion:
            fin_abs = inicio_abs + timedelta(seconds=duracion)
        else:
            fin_abs = inicio_abs

        if (fin_abs - inicio_abs).total_seconds() < 60:
            continue

        cat_raw = str(row[col_cat]).strip() if col_cat >= 0 and not pd.isna(row[col_cat]) else ""
        sitio_raw = str(row[col_sitio]).strip() if col_sitio >= 0 and not pd.isna(row[col_sitio]) else ""
        categoria = CatCategory.from_raw(cat_raw).guess_from_site(sitio_raw).nombre

        records.append(
            {
                "vehiculo": vehiculo,
                "inicio": inicio_abs,
                "fin": fin_abs,
                "categoria": categoria,
                "sitio": sitio_raw,
                "tipo": "VISITA",
            }
        )

    return pd.DataFrame.from_records(records)


def build_timeline(divisiones, visitas):
    pd = _require_pandas()
    eventos = []
    if divisiones is not None and len(divisiones):
        eventos.append(divisiones)
    if visitas is not None and len(visitas):
        visitas_norm = visitas.copy()
        visitas_norm["km"] = 0.0
        visitas_norm["minutos"] = (visitas_norm["fin"] - visitas_norm["inicio"]).dt.total_seconds() / 60.0
        visitas_norm["cliente_site"] = visitas_norm.get("sitio", "")
        visitas_norm["division"] = visitas_norm.get("division", "")
        visitas_norm["servicio_id"] = visitas_norm.apply(
            lambda row: servicio_id_from_components(row["vehiculo"], row["inicio"].date(), 0), axis=1
        )
        eventos.append(visitas_norm)

    if not eventos:
        return pd.DataFrame()

    timeline = pd.concat(eventos, ignore_index=True)
    timeline = timeline.sort_values(["vehiculo", "inicio"]).reset_index(drop=True)
    return timeline


__all__ = [
    "normalize",
    "find_header_row",
    "find_col_any_in_row",
    "date_only_ex2",
    "time_to_sec_ex",
    "normalize_vehiculo_key",
    "servicio_id_from_components",
    "find_header_row_visitas",
    "find_header_row_cargas",
    "extraer_division_desde_nombre",
    "read_division_excel",
    "read_visitas_excel",
    "build_timeline",
    "CatCategory",
]
