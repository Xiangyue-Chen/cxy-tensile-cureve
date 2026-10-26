from __future__ import annotations

import csv
from io import StringIO
import json
from pathlib import Path
import warnings

import pytest

from pandat_web import (
    PandatError,
    build_download_bundle,
    parse_pandat_line,
    render_phase_figure,
)


def _dat(rows: list[str], *, temperature_unit: str = "K") -> bytes:
    header = (
        "T\tphase_name\tf(@Alpha)\tf(@Liquid)\t"
        "x(Fe@Alpha)\tx(Ni@Alpha)\tx(Fe@Liquid)\tx(Ni@Liquid)"
    )
    units = f"{temperature_unit}\t\t\t\t\tmole/mole\tmole/mole\tmole/mole\tmole/mole"
    return ("\n".join([header, units, *rows]) + "\n").encode("utf-8")


def test_parses_kelvin_line_result_and_room_composition() -> None:
    result = parse_pandat_line(
        _dat(
            [
                "298.15\tAlpha\t1\t\t0.8\t0.2\t\t",
                "773.15\tAlpha+Liquid\t0.6\t0.4\t0.7\t0.3\t0.4\t0.6",
                "1273.15\tLiquid\t\t1\t\t\t0.3\t0.7",
            ]
        ),
        filename="demo.dat",
    )

    assert result.filename == "demo.dat"
    assert result.temperatures_C == pytest.approx((25.0, 500.0, 1000.0))
    assert result.phases == ("Alpha", "Liquid")
    assert result.fractions["Alpha"] == pytest.approx((1.0, 0.6, 0.0))
    assert result.room_compositions_at_percent["Alpha"] == pytest.approx(
        {"Fe": 80.0, "Ni": 20.0}
    )
    assert result.room_phase_fractions_percent == pytest.approx({"Alpha": 100.0})
    assert result.melting_start_C == pytest.approx(500.0)
    assert result.melting_end_C == pytest.approx(1000.0)


def test_detects_celsius_from_units() -> None:
    result = parse_pandat_line(
        _dat(
            [
                "25\tAlpha\t1\t\t1\t0\t\t",
                "900\tLiquid\t\t1\t\t\t1\t0",
            ],
            temperature_unit="C",
        ),
        filename="celsius.dat",
    )

    assert result.temperatures_C == pytest.approx((25.0, 900.0))


def test_collapses_duplicate_transition_temperature_without_zero_width_region() -> None:
    result = parse_pandat_line(
        _dat(
            [
                "298.15\tAlpha\t1\t\t1\t0\t\t",
                "773.15\tAlpha+Liquid\t0\t1\t1\t0\t0.5\t0.5",
                "773.15\tLiquid\t\t1\t\t\t0.5\t0.5",
                "1273.15\tLiquid\t\t1\t\t\t0.4\t0.6",
            ]
        ),
        filename="duplicate.dat",
    )

    assert result.temperatures_C == pytest.approx((25.0, 500.0, 1000.0))
    assert [(region.lower_C, region.upper_C, region.phases) for region in result.regions] == [
        (25.0, 500.0, ("Alpha",)),
        (500.0, 1000.0, ("Liquid",)),
    ]
    assert all(region.upper_C > region.lower_C for region in result.regions)


def test_phase_threshold_controls_trace_phase_regions() -> None:
    data = _dat(
        [
            "298.15\tAlpha+Liquid\t0.999999999\t0.000000001\t1\t0\t0.5\t0.5",
            "773.15\tAlpha+Liquid\t0.9\t0.1\t1\t0\t0.5\t0.5",
            "1273.15\tAlpha+Liquid\t0.5\t0.5\t1\t0\t0.5\t0.5",
        ]
    )

    result = parse_pandat_line(data, filename="trace.dat", phase_threshold=1e-8)

    assert result.regions[0].phases == ("Alpha",)
    assert result.regions[1].phases == ("Alpha", "Liquid")
    assert result.melting_start_C == pytest.approx(500.0)
    assert result.melting_end_C is None


@pytest.mark.parametrize(
    ("data", "message"),
    [
        (b"phase_name\tf(@Alpha)\n\t1\n", "温度列"),
        (b"T\tphase_name\nK\t\n298.15\tAlpha\n", "相分数列"),
        (b"T\tf(@Alpha)\nK\t\nnot-a-number\t1\n", "温度不是有效数字"),
        (b"", "空文件"),
    ],
)
def test_rejects_malformed_line_results(data: bytes, message: str) -> None:
    with pytest.raises(PandatError, match=message):
        parse_pandat_line(data, filename="bad.dat")


def test_rejects_invalid_threshold_and_oversized_upload() -> None:
    with pytest.raises(PandatError, match="阈值"):
        parse_pandat_line(_dat(["298.15\tAlpha\t1\t\t1\t0\t\t"]), filename="bad.dat", phase_threshold=0)
    with pytest.raises(PandatError, match="20 MiB"):
        parse_pandat_line(b"x" * (20 * 1024 * 1024 + 1), filename="large.dat")


def test_rejects_oversized_csv_field_with_chinese_error() -> None:
    data = (
        "T\tf(@Alpha)\n"
        "K\t\n"
        f"{'1' * 140_000}\t1\n"
    ).encode("utf-8")

    with pytest.raises(PandatError, match="制表符结构无法解析"):
        parse_pandat_line(data, filename="bad-field.dat")


def test_renders_reference_style_png_and_single_page_pdf() -> None:
    result = parse_pandat_line(
        _dat(
            [
                "298.15\tAlpha\t1\t\t0.8\t0.2\t\t",
                "773.15\tAlpha+Liquid\t0.6\t0.4\t0.7\t0.3\t0.4\t0.6",
                "1273.15\tLiquid\t\t1\t\t\t0.3\t0.7",
            ]
        ),
        filename="demo.dat",
    )

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        figure = render_phase_figure(result, sample_name="演示样品")

    assert figure.png.startswith(b"\x89PNG\r\n\x1a\n")
    assert len(figure.png) > 20_000
    assert figure.pdf.startswith(b"%PDF-")
    assert len(figure.pdf) > 5_000
    assert figure.pdf.count(b"/Type /Page") == 2  # one page plus the /Pages tree
    assert not [warning for warning in caught if "Glyph" in str(warning.message)]


def test_builds_csv_and_json_downloads_without_dropping_regions() -> None:
    result = parse_pandat_line(
        _dat(
            [
                "298.15\tAlpha\t1\t\t0.8\t0.2\t\t",
                "773.15\tAlpha+Liquid\t0.6\t0.4\t0.7\t0.3\t0.4\t0.6",
                "1273.15\tLiquid\t\t1\t\t\t0.3\t0.7",
                "1773.15\tLiquid\t\t1\t\t\t0.2\t0.8",
            ]
        ),
        filename="demo.dat",
    )

    files = build_download_bundle(result, sample_name="演示样品")

    assert set(files) == {
        "phase_diagram.png",
        "phase_diagram.pdf",
        "phase_fractions.csv",
        "phase_regions.csv",
        "phase_metadata.json",
    }
    fraction_rows = list(csv.DictReader(StringIO(files["phase_fractions.csv"].decode("utf-8-sig"))))
    region_rows = list(csv.DictReader(StringIO(files["phase_regions.csv"].decode("utf-8-sig"))))
    metadata = json.loads(files["phase_metadata.json"])
    assert len(fraction_rows) == len(result.temperatures_C)
    assert len(region_rows) == len(result.regions)
    assert metadata["sample_name"] == "演示样品"
    assert metadata["melting_interval_C"] == [500.0, 1000.0]
    assert metadata["phase_regions"][-1]["phases"] == ["Liquid"]


def test_l12_label_renders_without_missing_glyph_warning() -> None:
    data = _dat(
        [
            "298.15\tL12_FCC\t1\t\t0.7\t0.3\t\t",
            "773.15\tL12_FCC+Liquid\t0.5\t0.5\t0.6\t0.4\t0.4\t0.6",
            "1273.15\tLiquid\t\t1\t\t\t0.3\t0.7",
        ]
    ).replace(b"@Alpha", b"@L12_FCC")
    result = parse_pandat_line(data, filename="l12.dat")

    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always")
        render_phase_figure(result, sample_name="L12 test")

    assert not [warning for warning in caught if "Glyph" in str(warning.message)]


def test_trace_phase_below_region_threshold_is_still_plotted(monkeypatch) -> None:
    import matplotlib.axes

    data = (
        "T\tphase_name\tf(@A)\tf(@Trace)\n"
        "K\t\t\t\n"
        "298.15\tA+Trace\t0.999999999\t0.000000001\n"
        "773.15\tA\t1\t\n"
    ).encode("utf-8")
    result = parse_pandat_line(data, filename="trace.dat", phase_threshold=1e-8)
    labels: list[str] = []
    original_plot = matplotlib.axes.Axes.plot

    def capture_plot(self, *args, **kwargs):
        labels.append(kwargs.get("label"))
        return original_plot(self, *args, **kwargs)

    monkeypatch.setattr(matplotlib.axes.Axes, "plot", capture_plot)
    render_phase_figure(result, sample_name="Trace test")

    assert {label for label in labels if label} == {"A", "Trace"}


def test_sample_name_with_mathtext_characters_renders_literally() -> None:
    result = parse_pandat_line(
        _dat(
            [
                "298.15\tAlpha\t1\t\t0.8\t0.2\t\t",
                "773.15\tLiquid\t\t1\t\t\t0.4\t0.6",
            ]
        ),
        filename="literal-name.dat",
    )

    figure = render_phase_figure(result, sample_name=r"$\badcommand$ sample")

    assert figure.png.startswith(b"\x89PNG")


def test_c63_example_matches_validated_phase_regions_and_melting_interval() -> None:
    source = Path(__file__).with_name("examples") / "C63_C9_line_up.dat"

    result = parse_pandat_line(source.read_bytes(), filename=source.name)

    assert len(result.phases) == 8
    assert len(result.regions) == 12
    assert result.melting_start_C == pytest.approx(1179.375)
    assert result.melting_end_C == pytest.approx(1371.614727)
    assert sum(result.room_phase_fractions_percent.values()) == pytest.approx(100.0, abs=0.001)
