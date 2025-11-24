import datetime as dt
from pathlib import Path
import sys

import pytest

pd = pytest.importorskip("pandas")

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from analysis.cli import main
from analysis.normalization import (
    _build_division_records,
    _detect_division_header_row,
    _map_division_columns,
    _process_division_dataframe,
    _validate_division_rows,
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


def test_division_column_mapping(sample_division_file: Path):
    raw = pd.read_excel(sample_division_file, header=None)
    header = _detect_division_header_row(raw)
    columns = _map_division_columns(raw, header)

    assert columns == {
        "km": 1,
        "fecha": 2,
        "hora_ini": 3,
        "hora_fin": 4,
        "vehiculo": 0,
        "cliente": 5,
    }


def test_division_validation_filters_invalid_rows():
    raw = pd.DataFrame(
        [
            ["Vehículo", "Kilómetros", "Fecha Inicio", "Hora Inicio", "Hora Fin", "Cliente"],
            ["AA-01", 10, "15/01/2024", "07:30", "08:00", "Cliente A"],
            ["", 5, "", "07:00", "07:30", "Cliente B"],
        ]
    )
    header = _detect_division_header_row(raw)
    columns = _map_division_columns(raw, header)
    validated = _validate_division_rows(raw, header, columns, date_order="DMY")

    assert len(validated) == 1
    assert validated.loc[0, "vehiculo"] == "AA01"
    assert validated.loc[0, "hora_fin_sec"] > validated.loc[0, "hora_ini_sec"]


def test_division_record_builder_preserves_metadata(sample_division_file: Path):
    raw = pd.read_excel(sample_division_file, header=None)
    header = _detect_division_header_row(raw)
    columns = _map_division_columns(raw, header)
    validated = _validate_division_rows(raw, header, columns, date_order="MDY")

    records = _build_division_records(validated, path=str(sample_division_file))
    assert list(records.columns)[:4] == ["division", "vehiculo", "inicio", "fin"]
    assert records["division"].iloc[0] == "Division Norte"
    assert records["servicio_id"].iloc[0].endswith("-000")


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


def test_division_orchestration_matches_public_api(sample_division_file: Path):
    raw = pd.read_excel(sample_division_file, header=None)
    orchestrated = _process_division_dataframe(raw, path=str(sample_division_file), date_order="MDY")
    public_api = read_division_excel(sample_division_file, date_order="MDY")

    pd.testing.assert_frame_equal(
        orchestrated.reset_index(drop=True), public_api.reset_index(drop=True)
    )


def test_cli_requires_output(sample_division_file: Path, sample_visitas_file: Path, capsys):
    with pytest.raises(SystemExit) as excinfo:
        main(
            [
                "--division-file",
                str(sample_division_file),
                "--visitas-file",
                str(sample_visitas_file),
            ]
        )

    assert excinfo.value.code == 2
    stderr = capsys.readouterr().err
    assert "--output-csv o --output-excel" in stderr
