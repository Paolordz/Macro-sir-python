"""Command-line entry point replacing interactive VBA dialogs.

Usage example::

    python -m analysis.cli \
        --division-file /ruta/division.xlsx \
        --visitas-file /ruta/visitas.xlsx \
        --output-csv timeline.csv
"""
from __future__ import annotations

import argparse
import os
from pathlib import Path
from typing import Optional

from .normalization import build_timeline, read_division_excel, read_visitas_excel


def _env_or_default(name: str, default: Optional[str]) -> Optional[str]:
    value = os.environ.get(name)
    return value if value else default


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generar tabla de línea de tiempo/Gantt desde reportes Excel")
    parser.add_argument(
        "--division-file",
        dest="division_file",
        default=_env_or_default("DIVISION_FILE", None),
        help="Ruta al archivo de división (equivale al diálogo GetFolderPath/GetVisitasPath)",
    )
    parser.add_argument(
        "--visitas-file",
        dest="visitas_file",
        default=_env_or_default("VISITAS_FILE", None),
        help="Ruta al archivo Reporte de Visitas",
    )
    parser.add_argument("--output-csv", dest="output_csv", default=_env_or_default("OUTPUT_CSV", None))
    parser.add_argument("--output-excel", dest="output_excel", default=_env_or_default("OUTPUT_XLSX", None))
    parser.add_argument("--division-date-order", dest="division_date_order", default="MDY")
    parser.add_argument("--visitas-date-order", dest="visitas_date_order", default="DMY")
    parser.add_argument("--division-sheet", dest="division_sheet", default=None)
    parser.add_argument("--visitas-sheet", dest="visitas_sheet", default=None)
    return parser.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    if not args.division_file:
        raise SystemExit("--division-file o la variable de entorno DIVISION_FILE es obligatoria")
    if not args.visitas_file:
        raise SystemExit("--visitas-file o la variable de entorno VISITAS_FILE es obligatoria")

    div_df = read_division_excel(args.division_file, date_order=args.division_date_order, sheet_name=args.division_sheet)
    visitas_df = read_visitas_excel(args.visitas_file, date_order=args.visitas_date_order, sheet_name=args.visitas_sheet)
    timeline_df = build_timeline(div_df, visitas_df)

    if args.output_csv:
        Path(args.output_csv).parent.mkdir(parents=True, exist_ok=True)
        timeline_df.to_csv(args.output_csv, index=False)
    if args.output_excel:
        Path(args.output_excel).parent.mkdir(parents=True, exist_ok=True)
        timeline_df.to_excel(args.output_excel, index=False)

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
