import datetime as dt
from pathlib import Path

import pytest

pd = pytest.importorskip("pandas")

from analysis.normalization import (
    build_timeline,
    find_header_row_cargas,
    find_header_row_visitas,
    read_division_excel,
    read_visitas_excel,
)


@pytest.fixture()
def sample_division_file(tmp_path: Path) -> Path:
    data = pd.DataFrame(
        [
            ["Vehículo", "Kilómetros", "Fecha Inicio", "Hora Inicio", "Hora Fin", "Cliente"],
            ["AA-01", 10, "01/15/2024", "07:30", "09:00", "Cliente A"],
            ["BB-02", 5, "16/01/2024", "630", "700", "Cliente B"],
        ]
    )
    path = tmp_path / "Division Norte.xlsx"
    data.to_excel(path, header=False, index=False)
    return path


@pytest.fixture()
def sample_visitas_file(tmp_path: Path) -> Path:
    data = pd.DataFrame(
        [
            ["Unidad", "Fecha Llegada", "Hora Llegada", "Fecha Salida", "Hora Salida", "Categoría", "Sitio"],
            ["AA-01", "15/01/2024", "10:00", "15/01/2024", "11:00", "Patio", "Patio Central"],
            ["BB-02", "15/01/2024", "12:00", "15/01/2024", "13:30", "", "Taller Regional"],
        ]
    )
    path = tmp_path / "visitas.xlsx"
    data.to_excel(path, header=False, index=False)
    return path


def test_visitas_header_detection(sample_visitas_file: Path):
    raw = pd.read_excel(sample_visitas_file, header=None)
    header = find_header_row_visitas(raw)
    assert header == 0


def test_cargas_header_detection():
    df = pd.DataFrame(
        [
            ["otra", "fila"],
            ["Vehículo", "Fecha", "División"],
            ["AA-01", "01/01/2024", "DIV A"],
        ]
    )
    assert find_header_row_cargas(df) == 1


def test_read_division_and_visitas(sample_division_file: Path, sample_visitas_file: Path):
    div_df = read_division_excel(sample_division_file, date_order="MDY")
    visitas_df = read_visitas_excel(sample_visitas_file, date_order="DMY")

    assert set(div_df.columns) >= {"vehiculo", "inicio", "fin", "servicio_id"}
    assert set(visitas_df.columns) >= {"vehiculo", "inicio", "fin", "categoria"}

    timeline = build_timeline(div_df, visitas_df)
    assert len(timeline) == len(div_df) + len(visitas_df)
    assert list(timeline.sort_values("inicio").vehiculo.unique()) == ["AA01", "BB02"]
