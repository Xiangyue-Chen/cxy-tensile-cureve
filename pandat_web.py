"""Parse native Pandat line results for the Streamlit phase-diagram tool."""
from __future__ import annotations

from dataclasses import dataclass
import csv
from io import BytesIO, StringIO
import json
import math
import os
from pathlib import Path
import re


MAX_UPLOAD_BYTES = 20 * 1024 * 1024
_PHASE_RE = re.compile(r"^f\(@(.+)\)$")
_COMPOSITION_RE = re.compile(r"^x\(([^@()]+)@(.+)\)$")


class PandatError(ValueError):
    """A user-facing Pandat upload or analysis error."""


@dataclass(frozen=True)
class PhaseRegion:
    lower_C: float
    upper_C: float
    phases: tuple[str, ...]


@dataclass(frozen=True)
class PandatResult:
    filename: str
    temperatures_C: tuple[float, ...]
    phases: tuple[str, ...]
    fractions: dict[str, tuple[float, ...]]
    regions: tuple[PhaseRegion, ...]
    room_temperature_C: float
    room_phase_fractions_percent: dict[str, float]
    room_compositions_at_percent: dict[str, dict[str, float]]
    overall_composition_at_percent: dict[str, float]
    melting_start_C: float | None
    melting_end_C: float | None
    phase_threshold: float


@dataclass(frozen=True)
class FigureBundle:
    png: bytes
    pdf: bytes


def _decode(data: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-16", "gb18030"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    raise PandatError("文件编码无法识别，请导出为 UTF-8 制表符 DAT 文件。")


def _number(value: str, *, row_number: int, label: str, empty_zero: bool = False) -> float:
    text = value.strip()
    if not text and empty_zero:
        return 0.0
    try:
        number = float(text)
    except ValueError as exc:
        raise PandatError(f"第 {row_number} 行的{label}不是有效数字。") from exc
    if not math.isfinite(number):
        raise PandatError(f"第 {row_number} 行的{label}不是有限数字。")
    return number


def _is_celsius(unit: str, raw_temperatures: list[float]) -> bool:
    normalized = unit.strip().lower().replace("°", "").replace("℃", "c")
    if normalized in {"c", "degc", "celsius"}:
        return True
    if normalized in {"k", "kelvin"}:
        return False
    return min(raw_temperatures) < 200.0


def _phase_regions(
    temperatures: tuple[float, ...],
    phases: tuple[str, ...],
    fractions: dict[str, tuple[float, ...]],
    threshold: float,
) -> tuple[PhaseRegion, ...]:
    phase_sets = [tuple(phase for phase in phases if fractions[phase][i] > threshold) for i in range(len(temperatures))]
    groups: list[tuple[int, int, tuple[str, ...]]] = []
    start = 0
    current = phase_sets[0]
    for index, phase_set in enumerate(phase_sets[1:], start=1):
        if phase_set != current:
            groups.append((start, index, current))
            start = index
            current = phase_set
    groups.append((start, len(temperatures), current))

    regions: list[PhaseRegion] = []
    for index, (start_index, _stop_index, phase_set) in enumerate(groups):
        lower = temperatures[start_index]
        upper = temperatures[groups[index + 1][0]] if index + 1 < len(groups) else temperatures[-1]
        if upper > lower:
            regions.append(PhaseRegion(lower, upper, phase_set))
    return tuple(regions)


def parse_pandat_line(
    data: bytes,
    *,
    filename: str,
    phase_threshold: float = 1e-8,
) -> PandatResult:
    """Parse a native Pandat tab-delimited line result without changing curve values."""
    if not data:
        raise PandatError("上传的是空文件。")
    if len(data) > MAX_UPLOAD_BYTES:
        raise PandatError("单个 Pandat 文件不能超过 20 MiB。")
    if not math.isfinite(phase_threshold) or not 0 < phase_threshold < 1:
        raise PandatError("相存在阈值必须在 0 与 1 之间。")

    text = _decode(data)
    try:
        rows = list(csv.reader(StringIO(text), delimiter="\t"))
    except csv.Error as exc:
        raise PandatError(f"DAT 制表符结构无法解析：{exc}") from exc
    if not rows or not any(cell.strip() for cell in rows[0]):
        raise PandatError("上传的是空文件。")
    header = [cell.strip() for cell in rows[0]]
    if "T" not in header:
        raise PandatError("未找到 Pandat 温度列 T。")
    phase_columns = [(index, match.group(1)) for index, name in enumerate(header) if (match := _PHASE_RE.fullmatch(name))]
    if not phase_columns:
        raise PandatError("未找到 f(@相名) 相分数列。")
    phases = tuple(phase for _index, phase in phase_columns)
    temperature_index = header.index("T")
    unit_row = rows[1] if len(rows) > 1 else []
    temperature_unit = unit_row[temperature_index] if temperature_index < len(unit_row) else ""

    composition_columns: list[tuple[int, str, str]] = []
    overall_columns: list[tuple[int, str]] = []
    for index, name in enumerate(header):
        if match := _COMPOSITION_RE.fullmatch(name):
            composition_columns.append((index, match.group(2), match.group(1)))
        elif match := re.fullmatch(r"x\(([^@()]+)\)", name):
            overall_columns.append((index, match.group(1)))

    parsed_rows: list[dict[str, object]] = []
    start_index = 2 if len(rows) > 1 else 1
    for source_index, source_row in enumerate(rows[start_index:], start=start_index + 1):
        if not any(cell.strip() for cell in source_row):
            continue
        padded = source_row + [""] * max(0, len(header) - len(source_row))
        temperature = _number(padded[temperature_index], row_number=source_index, label="温度")
        fraction_values: dict[str, float] = {}
        for column_index, phase in phase_columns:
            value = _number(padded[column_index], row_number=source_index, label=f"{phase} 相分数", empty_zero=True)
            if value < -1e-9 or value > 1.0 + 1e-6:
                raise PandatError(f"第 {source_index} 行的 {phase} 相分数超出 0–1 范围。")
            fraction_values[phase] = min(1.0, max(0.0, value))
        compositions: dict[str, dict[str, float]] = {}
        for column_index, phase, element in composition_columns:
            value = padded[column_index].strip()
            if value:
                compositions.setdefault(phase, {})[element] = _number(
                    value, row_number=source_index, label=f"{phase} 中 {element} 成分"
                ) * 100.0
        overall = {
            element: _number(padded[column_index], row_number=source_index, label=f"总体 {element} 成分") * 100.0
            for column_index, element in overall_columns
            if padded[column_index].strip()
        }
        parsed_rows.append(
            {
                "temperature": temperature,
                "fractions": fraction_values,
                "compositions": compositions,
                "overall": overall,
            }
        )
    if not parsed_rows:
        raise PandatError("文件中没有可用的温度数据行。")

    raw_temperatures = [float(row["temperature"]) for row in parsed_rows]
    celsius = _is_celsius(temperature_unit, raw_temperatures)
    by_temperature: dict[float, dict[str, object]] = {}
    for row in parsed_rows:
        raw_temperature = float(row["temperature"])
        temperature_C = raw_temperature if celsius else raw_temperature - 273.15
        row["temperature_C"] = round(temperature_C, 10)
        by_temperature[row["temperature_C"]] = row
    ordered = [by_temperature[key] for key in sorted(by_temperature)]
    temperatures_C = tuple(float(row["temperature_C"]) for row in ordered)
    if len(temperatures_C) < 2 or temperatures_C[-1] <= temperatures_C[0]:
        raise PandatError("至少需要两个不同温度点。")
    fractions = {
        phase: tuple(float(row["fractions"][phase]) for row in ordered)  # type: ignore[index]
        for phase in phases
    }
    regions = _phase_regions(temperatures_C, phases, fractions, phase_threshold)
    room_index = min(range(len(temperatures_C)), key=lambda index: abs(temperatures_C[index] - 25.0))
    room_row = ordered[room_index]
    room_fractions = {
        phase: fractions[phase][room_index] * 100.0
        for phase in phases
        if fractions[phase][room_index] > phase_threshold
    }
    room_compositions = {
        phase: dict(elements)
        for phase, elements in room_row["compositions"].items()  # type: ignore[union-attr]
        if phase in room_fractions
    }
    liquid_phase = next((phase for phase in phases if phase.lower() == "liquid"), None)
    melting_start = None
    melting_end = None
    if liquid_phase:
        for temperature, value in zip(temperatures_C, fractions[liquid_phase]):
            if value > phase_threshold:
                melting_start = temperature
                break
        for temperature, value in zip(temperatures_C, fractions[liquid_phase]):
            if value >= 1.0 - phase_threshold:
                melting_end = temperature
                break

    return PandatResult(
        filename=filename,
        temperatures_C=temperatures_C,
        phases=phases,
        fractions=fractions,
        regions=regions,
        room_temperature_C=temperatures_C[room_index],
        room_phase_fractions_percent=room_fractions,
        room_compositions_at_percent=room_compositions,
        overall_composition_at_percent=dict(room_row["overall"]),  # type: ignore[arg-type]
        melting_start_C=melting_start,
        melting_end_C=melting_end,
        phase_threshold=phase_threshold,
    )


_COLORS = (
    "#0072B2",
    "#E69F00",
    "#009E73",
    "#D55E00",
    "#9467BD",
    "#4D4D4D",
    "#8C564B",
    "#CC79A7",
    "#56B4E9",
    "#F0E442",
)
_LINESTYLES = ("-", "--", "-.", ":")
_KNOWN_COLORS = {
    "Bcc_B2": "#0072B2",
    "Bcc": "#9467BD",
    "Ni3Ti": "#E69F00",
    "Fcc": "#D55E00",
    "L12_FCC": "#009E73",
    "MB2_hP3": "#4D4D4D",
    "Cr5B3": "#8C564B",
    "Liquid": "#CC79A7",
}
_KNOWN_STYLES = {
    "Bcc_B2": "-",
    "Bcc": "--",
    "Ni3Ti": "-",
    "Fcc": "-",
    "L12_FCC": "-.",
    "MB2_hP3": "--",
    "Cr5B3": ":",
    "Liquid": ":",
}


def _phase_label(phase: str) -> str:
    return {
        "Fcc": "FCC",
        "Bcc": "BCC",
        "L12_FCC": "L1$_2$",
        "MB2_hP3": "MB2",
        "Liquid": "液相",
    }.get(phase, phase)


def _plain_phase_label(phase: str) -> str:
    return {
        "Fcc": "FCC",
        "Bcc": "BCC",
        "L12_FCC": "L1-2",
        "MB2_hP3": "MB2",
        "Liquid": "液相",
    }.get(phase, phase)


def _font_family(font_manager) -> str:
    candidates = [
        os.environ.get("PANDAT_FONT_PATH", ""),
        "C:/Windows/Fonts/msyh.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            font_manager.fontManager.addfont(candidate)
            return font_manager.FontProperties(fname=candidate).get_name()
    return "DejaVu Sans"


def _region_text(phases: tuple[str, ...]) -> str:
    return " + ".join(_plain_phase_label(phase) for phase in phases) or "无稳定相"


def render_phase_figure(result: PandatResult, *, sample_name: str) -> FigureBundle:
    """Render separate native phase curves with room notes and a phase-region band."""
    import matplotlib

    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    from matplotlib import font_manager
    from matplotlib.patches import ConnectionPatch, Rectangle

    font_family = _font_family(font_manager)
    with plt.rc_context(
        {
            "font.family": font_family,
            "axes.unicode_minus": False,
            "font.size": 10,
            "pdf.fonttype": 42,
            "ps.fonttype": 42,
        }
    ):
        height = 8.6 if len(result.regions) > 8 else 7.8
        fig = plt.figure(figsize=(12.4, height), layout="constrained")
        grid = fig.add_gridspec(
            2,
            2,
            width_ratios=[3.3, 1.25],
            height_ratios=[4.7, 1.8 if len(result.regions) > 8 else 1.3],
            hspace=0.035,
        )
        ax = fig.add_subplot(grid[0, 0])
        notes = fig.add_subplot(grid[0, 1])
        notes.axis("off")

        phase_colors = {
            phase: _KNOWN_COLORS.get(phase, _COLORS[index % len(_COLORS)])
            for index, phase in enumerate(result.phases)
        }
        plotted = list(result.phases)
        for index, phase in enumerate(plotted):
            ax.plot(
                result.temperatures_C,
                [value * 100.0 for value in result.fractions[phase]],
                color=phase_colors[phase],
                label=_phase_label(phase),
                lw=1.75,
                linestyle=_KNOWN_STYLES.get(phase, _LINESTYLES[index % len(_LINESTYLES)]),
            )
        minimum, maximum = result.temperatures_C[0], result.temperatures_C[-1]
        ax.set(
            xlim=(minimum, maximum),
            ylim=(-1, 103),
            xlabel="温度（℃）",
            ylabel="相摩尔分数（%）",
        )
        ax.grid(alpha=0.18)
        ax.text(
            0.022,
            0.98,
            sample_name.strip() or Path(result.filename).stem,
            transform=ax.transAxes,
            ha="left",
            va="top",
            fontsize=15,
            weight="bold",
            parse_math=False,
        )
        legend = ax.legend(
            loc="upper left",
            bbox_to_anchor=(0.012, 0.925),
            fontsize=8.3,
            ncol=2,
            framealpha=0.9,
        )
        for text, phase in zip(legend.get_texts(), plotted):
            if phase not in _KNOWN_COLORS:
                text.set_parse_math(False)

        band = fig.add_subplot(grid[1, 0], sharex=ax)
        band.set(ylim=(0, 1.05), yticks=[])
        band.spines[["left", "right", "top", "bottom"]].set_visible(False)
        band.tick_params(axis="x", bottom=False, labelbottom=False)
        total_range = maximum - minimum
        narrow: list[tuple[int, PhaseRegion]] = []
        for index, region in enumerate(result.regions):
            width = region.upper_C - region.lower_C
            fill = _COLORS[index % len(_COLORS)] + "33"
            band.add_patch(
                Rectangle(
                    (region.lower_C, 0.82),
                    width,
                    0.18,
                    facecolor=fill,
                    edgecolor="#777777",
                    lw=0.65,
                )
            )
            if width / total_range >= 0.09:
                band.text(
                    (region.lower_C + region.upper_C) / 2.0,
                    0.91,
                    _region_text(region.phases),
                    ha="center",
                    va="center",
                    fontsize=7.2,
                    color="#333333",
                    parse_math=False,
                )
            else:
                narrow.append((index, region))
        for region in result.regions[1:]:
            temperature = region.lower_C
            ax.axvline(temperature, color="#A5A5A5", linestyle=":", lw=0.7, alpha=0.65, zorder=0)
            fig.add_artist(
                ConnectionPatch(
                    xyA=(temperature, -1),
                    coordsA=ax.transData,
                    xyB=(temperature, 1),
                    coordsB=band.transData,
                    color="#A5A5A5",
                    linestyle=":",
                    lw=0.7,
                    alpha=0.65,
                )
            )
        if narrow:
            positions = [minimum + total_range * (index + 1) / (len(narrow) + 1) for index in range(len(narrow))]
            for callout_index, ((_region_index, region), text_x) in enumerate(zip(narrow, positions)):
                text_y = 0.47 if callout_index % 2 == 0 else 0.12
                label = (
                    f"{_region_text(region.phases)}\n"
                    f"{region.lower_C:.0f}–{region.upper_C:.0f} ℃"
                )
                band.annotate(
                    label,
                    xy=((region.lower_C + region.upper_C) / 2.0, 0.82),
                    xytext=(text_x, text_y),
                    ha="center",
                    va="center",
                    fontsize=6.8,
                    color="#333333",
                    linespacing=1.2,
                    annotation_clip=False,
                    arrowprops=dict(arrowstyle="-", color="#666666", lw=0.75),
                )

        notes.text(0.04, 0.94, f"{result.room_temperature_C:.0f} ℃时的成分特点", fontsize=12, weight="bold", transform=notes.transAxes)
        major_phases = sorted(
            result.room_phase_fractions_percent,
            key=result.room_phase_fractions_percent.get,
            reverse=True,
        )[:3]
        y_positions = (0.78, 0.59, 0.40)
        for y, phase in zip(y_positions, major_phases):
            fraction = result.room_phase_fractions_percent[phase]
            composition = result.room_compositions_at_percent.get(phase, {})
            main_elements = sorted(composition.items(), key=lambda item: item[1], reverse=True)
            details = "、".join(f"{element} {value:.1f}%" for element, value in main_elements if value >= 0.1)
            notes.text(
                0.04,
                y,
                f"{fraction:.1f}% {_plain_phase_label(phase)}",
                fontsize=10.7,
                weight="bold",
                color=phase_colors[phase],
                transform=notes.transAxes,
                parse_math=False,
            )
            notes.text(
                0.04,
                y - 0.035,
                f"相内成分：{details or '无可用数据'}",
                fontsize=8.6,
                color="#444444",
                va="top",
                wrap=True,
                transform=notes.transAxes,
            )
        trace_phases = [
            f"{fraction:.2f}% {_plain_phase_label(phase)}"
            for phase, fraction in result.room_phase_fractions_percent.items()
            if phase not in major_phases
        ]
        if trace_phases:
            notes.text(0.04, 0.25, "微量相：" + "；".join(trace_phases), fontsize=8.2, color="#555555", wrap=True, transform=notes.transAxes)
        if result.melting_start_C is not None and result.melting_end_C is not None:
            melting_text = f"熔化区间：{result.melting_start_C:.0f}–{result.melting_end_C:.0f} ℃"
        elif result.melting_start_C is not None:
            melting_text = f"开始熔化：{result.melting_start_C:.0f} ℃；未达到完全液化"
        else:
            melting_text = "上传温区内未检测到液相"
        notes.text(0.04, 0.15, melting_text, fontsize=10, transform=notes.transAxes)
        notes.text(
            0.04,
            0.015,
            f"成分标注对应最接近25 ℃的数据点（{result.room_temperature_C:.1f} ℃），\n不代表各温度下成分保持不变。",
            fontsize=8.7,
            color="#666666",
            linespacing=1.6,
            transform=notes.transAxes,
        )

        png_stream = BytesIO()
        pdf_stream = BytesIO()
        try:
            fig.savefig(png_stream, format="png", dpi=180)
            fig.savefig(pdf_stream, format="pdf")
        except (RuntimeError, ValueError) as exc:
            raise PandatError(f"绘图失败，请检查样品名和相名称：{exc}") from exc
        finally:
            plt.close(fig)
    return FigureBundle(png=png_stream.getvalue(), pdf=pdf_stream.getvalue())


def _csv_bytes(rows: list[dict[str, object]], fieldnames: list[str]) -> bytes:
    stream = StringIO(newline="")
    writer = csv.DictWriter(stream, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(rows)
    return stream.getvalue().encode("utf-8-sig")


def build_download_bundle(result: PandatResult, *, sample_name: str) -> dict[str, bytes]:
    """Build the plot and inspectable data tables for one parsed upload."""
    figure = render_phase_figure(result, sample_name=sample_name)
    fraction_rows = [
        {
            "Temperature_C": temperature,
            **{phase: result.fractions[phase][index] * 100.0 for phase in result.phases},
        }
        for index, temperature in enumerate(result.temperatures_C)
    ]
    region_rows = [
        {
            "lower_C": region.lower_C,
            "upper_C": region.upper_C,
            "phases": "+".join(region.phases),
        }
        for region in result.regions
    ]
    metadata = {
        "sample_name": sample_name,
        "source_filename": result.filename,
        "temperature_range_C": [result.temperatures_C[0], result.temperatures_C[-1]],
        "phase_threshold": result.phase_threshold,
        "phases": list(result.phases),
        "melting_interval_C": [result.melting_start_C, result.melting_end_C],
        "room_temperature_C": result.room_temperature_C,
        "room_phase_fractions_percent": result.room_phase_fractions_percent,
        "room_phase_compositions_at_percent": result.room_compositions_at_percent,
        "overall_composition_at_percent": result.overall_composition_at_percent,
        "phase_regions": [
            {
                "lower_C": region.lower_C,
                "upper_C": region.upper_C,
                "phases": list(region.phases),
            }
            for region in result.regions
        ],
        "evidence_note": "Pandat equilibrium line-result rendering; not experimental validation.",
    }
    return {
        "phase_diagram.png": figure.png,
        "phase_diagram.pdf": figure.pdf,
        "phase_fractions.csv": _csv_bytes(fraction_rows, ["Temperature_C", *result.phases]),
        "phase_regions.csv": _csv_bytes(region_rows, ["lower_C", "upper_C", "phases"]),
        "phase_metadata.json": json.dumps(metadata, ensure_ascii=False, indent=2).encode("utf-8"),
    }
