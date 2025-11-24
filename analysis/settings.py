"""Configuración centralizada para normalización y lectura de reportes.

Este módulo define constantes documentadas en un ``Settings`` dataclass para
los principales parámetros de búsqueda y reglas de fechas:

* Límites de filas a inspeccionar al buscar encabezados.
* Fila predeterminada a asumir cuando la heurística no encuentra resultados.
* Ajuste de año para valores con dos dígitos (``y < 100``) usados en
  ``date_only_ex2``.

Los valores toman como base las cifras históricas de los macros originales y
pueden sobrescribirse mediante variables de entorno o flags CLI. Los nombres de
variables de entorno se documentan en cada campo. Los flags CLI usan la misma
nomenclatura en kebab-case.
"""
from __future__ import annotations

from dataclasses import dataclass, replace
import os
from typing import Callable, Optional


def _int_from_env(name: str, default: int, *, parser: Callable[[str], int] = int) -> int:
    value = os.environ.get(name)
    if value is None:
        return default
    try:
        return parser(value)
    except ValueError:
        return default


@dataclass(frozen=True)
class Settings:
    """Agrupa los parámetros ajustables del módulo de análisis."""

    #: Número máximo de filas a inspeccionar al detectar encabezados genéricos.
    #: Variable de entorno: ``HEADER_MAX_ROWS``.
    header_max_rows: int = 25

    #: Fila predeterminada (0-based) para encabezados genéricos si la búsqueda
    #: no encuentra coincidencias. Variable de entorno: ``HEADER_DEFAULT_ROW``.
    header_default_row: int = 5

    #: Número máximo de filas a revisar en reportes de visitas.
    #: Variable de entorno: ``VISITAS_MAX_ROWS``.
    visitas_max_rows: int = 100

    #: Fila predeterminada (0-based) para encabezados de visitas.
    #: Variable de entorno: ``VISITAS_DEFAULT_ROW``.
    visitas_default_row: int = 7

    #: Número máximo de filas a revisar en reportes de cargas/servicios.
    #: Variable de entorno: ``CARGAS_MAX_ROWS``.
    cargas_max_rows: int = 100

    #: Fila predeterminada (0-based) para encabezados de cargas/servicios.
    #: Variable de entorno: ``CARGAS_DEFAULT_ROW``.
    cargas_default_row: int = 0

    #: Año base a sumar para valores de dos dígitos (``y < 100``).
    #: Variable de entorno: ``TWO_DIGIT_YEAR_BASE``.
    two_digit_year_base: int = 2000

    @classmethod
    def from_env(cls, env: Optional[dict[str, str]] = None) -> "Settings":
        """Construye la configuración leyendo variables de entorno.

        Si alguna variable no está presente o no es numérica se usa el valor
        por defecto definido en la clase.
        """

        environ = os.environ if env is None else env
        return cls(
            header_max_rows=_int_from_env("HEADER_MAX_ROWS", cls.header_max_rows, parser=int),
            header_default_row=_int_from_env("HEADER_DEFAULT_ROW", cls.header_default_row, parser=int),
            visitas_max_rows=_int_from_env("VISITAS_MAX_ROWS", cls.visitas_max_rows, parser=int),
            visitas_default_row=_int_from_env("VISITAS_DEFAULT_ROW", cls.visitas_default_row, parser=int),
            cargas_max_rows=_int_from_env("CARGAS_MAX_ROWS", cls.cargas_max_rows, parser=int),
            cargas_default_row=_int_from_env("CARGAS_DEFAULT_ROW", cls.cargas_default_row, parser=int),
            two_digit_year_base=_int_from_env("TWO_DIGIT_YEAR_BASE", cls.two_digit_year_base, parser=int),
        )

    def override_with_cli(self, args: object) -> "Settings":
        """Devuelve una copia con sobrescrituras provenientes de argumentos CLI.

        Los atributos se leen solo si el objeto ``args`` expone el nombre del
        campo; si el valor es ``None`` se mantiene el valor actual.
        """

        updates = {}
        for field in (
            "header_max_rows",
            "header_default_row",
            "visitas_max_rows",
            "visitas_default_row",
            "cargas_max_rows",
            "cargas_default_row",
            "two_digit_year_base",
        ):
            if hasattr(args, field):
                value = getattr(args, field)
                if value is not None:
                    updates[field] = int(value)
        return replace(self, **updates) if updates else self


__all__ = ["Settings"]
