"""Helpers for converting legacy Excel macros into Python utilities."""

from .normalization import (
    CatCategory,
    date_only_ex2,
    extraer_division_desde_nombre,
    find_col_any_in_row,
    find_header_row,
    find_header_row_cargas,
    find_header_row_visitas,
    normalize,
    normalize_vehiculo_key,
    read_division_excel,
    read_visitas_excel,
    build_timeline,
    servicio_id_from_components,
    time_to_sec_ex,
)

__all__ = [
    "CatCategory",
    "date_only_ex2",
    "extraer_division_desde_nombre",
    "find_col_any_in_row",
    "find_header_row",
    "find_header_row_cargas",
    "find_header_row_visitas",
    "normalize",
    "normalize_vehiculo_key",
    "read_division_excel",
    "read_visitas_excel",
    "build_timeline",
    "servicio_id_from_components",
    "time_to_sec_ex",
]
