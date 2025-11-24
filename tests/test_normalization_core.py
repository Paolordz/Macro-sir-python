import datetime as dt

import pytest

from analysis.normalization import (
    CatCategory,
    date_only_ex2,
    normalize,
    normalize_vehiculo_key,
    servicio_id_from_components,
    time_to_sec_ex,
)


def test_normalize_removes_accents_and_spaces():
    assert normalize("  Fécha Inicio  ") == "fechainicio"
    assert normalize("Vehículo-1") == "vehiculo1"


def test_date_only_ex2_handles_mdy_and_dmy():
    assert date_only_ex2("12/31/2023", "MDY") == dt.date(2023, 12, 31)
    assert date_only_ex2("31-12-23", "DMY") == dt.date(2023, 12, 31)
    assert date_only_ex2("", "DMY") is None


@pytest.mark.parametrize(
    "value, order",
    [
        ("31/13/2024", "DMY"),
        ("2024-31-12", "MDY"),
        ("no-date", "DMY"),
        (0, "DMY"),
        (-5, "MDY"),
        (0.5, "MDY"),
        ("31/12/2024", "YMD"),
    ],
)
def test_date_only_ex2_rejects_invalid_inputs(value, order):
    assert date_only_ex2(value, order) is None


def test_time_to_sec_ex_accepts_fraction_and_text():
    assert time_to_sec_ex(0.5) == 43200
    assert time_to_sec_ex("12:30") == 12 * 3600 + 30 * 60
    assert time_to_sec_ex("0730") == 7 * 3600 + 30 * 60
    assert time_to_sec_ex("99:99") == 0


@pytest.mark.parametrize("value", ["48:00", "1260", "ab:cd", "-0100", None, ""])  # type: ignore[arg-type]
def test_time_to_sec_ex_rejects_out_of_range_and_bad_formats(value):
    assert time_to_sec_ex(value) == 0


def test_normalize_vehiculo_key_filters_chars():
    assert normalize_vehiculo_key("abc-123") == "ABC123"
    assert normalize_vehiculo_key(None) == ""


def test_servicio_id_matches_vba_format():
    fecha = dt.date(2024, 1, 1)
    assert servicio_id_from_components("Veh-01", fecha, 7) == "VEH01-20240101-007"
    # default NA when fecha is invalid
    assert servicio_id_from_components("", "n/a", -1).startswith("NA-00000000-")


def test_cat_category_guesses_from_site():
    assert CatCategory.from_raw("Patio Central").nombre == "Patio"
    fallback = CatCategory.from_raw("otros")
    assert fallback.guess_from_site("Taller Norte").nombre == "Taller"
