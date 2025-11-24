import datetime as dt
from pathlib import Path
import sys

import pytest

pd = pytest.importorskip("pandas")

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from analysis.cli import main
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


@pytest.mark.parametrize(
    "rows, finder, expected",
    [
        (
            [["foo", "bar"], ["x", "y"]],
            find_header_row_cargas,
            0,
        ),
        (
            [["Unidad"], ["algo"], [""]],
            find_header_row_visitas,
            7,
        ),
    ],
)
def test_header_detection_falls_back_to_default(rows, finder, expected):
    df = pd.DataFrame(rows)
    assert finder(df) == expected


def test_read_division_and_visitas(sample_division_file: Path, sample_visitas_file: Path):
    div_df = read_division_excel(sample_division_file, date_order="MDY")
    visitas_df = read_visitas_excel(sample_visitas_file, date_order="DMY")

    assert set(div_df.columns) >= {"vehiculo", "inicio", "fin", "servicio_id"}
    assert set(visitas_df.columns) >= {"vehiculo", "inicio", "fin", "categoria"}

    timeline = build_timeline(div_df, visitas_df)
    assert len(timeline) == len(div_df) + len(visitas_df)
    assert list(timeline.sort_values("inicio").vehiculo.unique()) == ["AA01", "BB02"]


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


def test_read_excel_raises_on_corrupted_files(tmp_path: Path):
    corrupted = tmp_path / "corrupted.xlsx"
    corrupted.write_bytes(b"not a valid excel file")

    with pytest.raises(Exception):
        read_division_excel(corrupted)


def test_read_excel_rejects_weird_encoding(tmp_path: Path):
    weird = tmp_path / "weird.xlsx"
    weird.write_text("ñandú", encoding="utf-16")

    with pytest.raises(Exception):
        read_visitas_excel(weird)


def test_cli_writes_outputs_and_uses_date_orders(sample_division_file: Path, sample_visitas_file: Path, tmp_path: Path):
    csv_out = tmp_path / "timeline.csv"
    xlsx_out = tmp_path / "timeline.xlsx"

    exit_code = main(
        [
            "--division-file",
            str(sample_division_file),
            "--visitas-file",
            str(sample_visitas_file),
            "--output-csv",
            str(csv_out),
            "--output-excel",
            str(xlsx_out),
            "--division-date-order",
            "MDY",
            "--visitas-date-order",
            "DMY",
        ]
    )

    assert exit_code == 0
    assert csv_out.exists()
    assert xlsx_out.exists()

    csv_df = pd.read_csv(csv_out)
    assert {"vehiculo", "inicio", "fin"}.issubset(csv_df.columns)


@pytest.mark.parametrize(
    "argv, message",
    [
        (["--visitas-file", "dummy", "--output-csv", "out.csv"], "division-file"),
        (["--division-file", "dummy", "--output-csv", "out.csv"], "visitas-file"),
    ],
)
def test_cli_fails_when_inputs_missing(argv, message):
    with pytest.raises(SystemExit) as excinfo:
        main(argv)

    assert message in str(excinfo.value)
