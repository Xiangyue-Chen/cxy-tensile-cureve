"""Streamlit front end for the isolated tensile-curve processor."""
from __future__ import annotations

import hashlib
from io import BytesIO
from pathlib import Path, PurePosixPath
import zipfile

import streamlit as st

from pandat_web import PandatError, PandatResult, build_download_bundle, parse_pandat_line
from tensile_web import (
    ProcessingError,
    ProcessingResult,
    Upload,
    build_plot_outputs,
    make_demo_uploads,
    process_uploads,
)


st.set_page_config(
    page_title="拉伸曲线在线处理",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown(
    """
    <style>
    :root { --ink:#123047; --muted:#5f7382; --accent:#0f8b8d; --paper:#f7fafb; }
    .stApp { background: linear-gradient(180deg, #f4faf9 0, #ffffff 28rem); color: var(--ink); }
    .block-container { max-width: 1180px; padding-top: 2.2rem; padding-bottom: 4rem; }
    .hero { border: 1px solid #d8e8e5; border-radius: 20px; padding: 2rem 2.2rem;
            background: rgba(255,255,255,.92); box-shadow: 0 16px 42px rgba(18,48,71,.08); }
    .eyebrow { color: var(--accent); font-weight: 750; letter-spacing:.08em; font-size:.8rem; }
    .hero h1 { color:var(--ink); margin:.35rem 0 .6rem; font-size:2.45rem; letter-spacing:-.04em; }
    .hero p { color:var(--muted); margin:0; font-size:1.02rem; max-width:780px; }
    .step { color:var(--accent); font-size:.78rem; font-weight:750; text-transform:uppercase; letter-spacing:.08em; }
    div[data-testid="stFileUploader"] { border: 1px dashed #8fbeba; border-radius: 16px; padding:.4rem; background:white; }
    div[data-testid="stMetric"] { background:white; border:1px solid #e1ecea; border-radius:14px; padding:1rem; }
    .notice { padding:.75rem 1rem; border-radius:12px; background:#eef7f6; color:#315b60; margin:.5rem 0 1rem; }
    </style>
    """,
    unsafe_allow_html=True,
)


def _uploaded_bytes(uploaded_files) -> list[Upload]:
    return [Upload(file.name, file.getvalue()) for file in uploaded_files]


def _fingerprint(uploaded_files) -> str:
    digest = hashlib.sha256()
    for uploaded in uploaded_files:
        data = uploaded.getvalue()
        digest.update(uploaded.name.encode("utf-8", errors="replace"))
        digest.update(len(data).to_bytes(8, "big"))
        digest.update(hashlib.sha256(data).digest())
    return digest.hexdigest()


def _clear_result() -> None:
    st.session_state.result = None
    st.session_state.result_source = None
    for key in ("plot_curve_selection", "plot_mode", "mark_performance_points", "plot_title", "custom_plot_preview", "preview_image"):
        st.session_state.pop(key, None)


def _run(uploads: list[Upload], source: str) -> None:
    try:
        for key in ("plot_curve_selection", "plot_mode", "mark_performance_points", "plot_title", "custom_plot_preview", "preview_image"):
            st.session_state.pop(key, None)
        with st.spinner("正在识别曲线、计算性能并生成结果…"):
            st.session_state.result = process_uploads(uploads)
        st.session_state.result_source = source
    except ProcessingError as exc:
        _clear_result()
        st.error(str(exc))


def _curve_label(curve: dict[str, object]) -> str:
    name = str(curve.get("sample_id") or curve.get("run_name") or curve.get("curve_id") or "未命名曲线")
    source = PurePosixPath(str(curve.get("source_file") or "")).name
    return f"{name} · {source}" if source else name


def _zip_outputs(files: dict[str, bytes]) -> bytes:
    stream = BytesIO()
    with zipfile.ZipFile(stream, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in files.items():
            archive.writestr(name, data)
    return stream.getvalue()


def _show_custom_plotter(result: ProcessingResult) -> None:
    st.subheader("Nature 风格自定义绘图")
    st.caption("自由选择一条或多条曲线；默认叠加到同一张图，也可分别生成。图片为白底、简洁坐标轴、色盲友好配色和 300 dpi PNG。")
    curve_map = {str(curve["curve_id"]): curve for curve in result.plot_curves}
    curve_ids = list(curve_map)
    selected_ids = st.multiselect(
        "选择曲线（可多选）",
        curve_ids,
        default=curve_ids,
        format_func=lambda curve_id: _curve_label(curve_map[curve_id]),
        key="plot_curve_selection",
    )
    mode_label = st.radio(
        "出图方式",
        ["叠加对比", "单独出图"],
        horizontal=True,
        key="plot_mode",
    )
    mark_performance_points = st.checkbox(
        "标注性能点（Rp0.2、Rm、Ag、At）",
        value=True,
        key="mark_performance_points",
        help="缺失或自动判据未通过的指标不会标注。Ag 与 Rm 对应同一峰值位置，Ag 以横轴三角形和引导线表示。",
    )
    title = st.text_input(
        "图标题（留空则不显示）",
        value="Tensile stress–strain curves",
        max_chars=120,
        key="plot_title",
    )
    if not selected_ids:
        st.info("请至少选择一条曲线。")
        return
    try:
        outputs = build_plot_outputs(
            [curve_map[curve_id] for curve_id in selected_ids],
            mode="overlay" if mode_label == "叠加对比" else "separate",
            title=title.strip(),
            mark_performance_points=mark_performance_points,
        )
    except ProcessingError as exc:
        st.error(str(exc))
        return

    if mode_label == "叠加对比":
        name, data = next(iter(outputs.items()))
        st.image(data, caption=f"已叠加 {len(selected_ids)} 条曲线", width="stretch")
        st.download_button(
            "下载所选曲线合并图 PNG",
            data=data,
            file_name=name,
            mime="image/png",
            type="primary",
            key="download_custom_overlay",
        )
        return

    preview_name = st.selectbox(
        "预览单独图片",
        list(outputs),
        key="custom_plot_preview",
    )
    st.image(outputs[preview_name], caption=preview_name, width="stretch")
    st.download_button(
        "下载所选曲线单独图片 ZIP",
        data=_zip_outputs(outputs),
        file_name="selected_curves_separate.zip",
        mime="application/zip",
        type="primary",
        key="download_custom_separate",
    )
    with st.expander("下载单张自定义图片"):
        for index, (name, data) in enumerate(outputs.items()):
            st.download_button(
                f"下载 {name}", data=data, file_name=name, mime="image/png",
                key=f"download_custom_single_{index}",
            )


def _show_results(result: ProcessingResult) -> None:
    st.success(f"处理完成：识别到 {len(result.rows)} 条曲线。")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("上传项", result.uploaded_count)
    col2.metric("有效曲线", len(result.rows))
    col3.metric("生成图片", len([name for name in result.files if name.lower().endswith(".png")]))
    col4.metric("文件级错误", len(result.errors))

    for message in result.warnings:
        st.warning(message)
    if result.errors:
        with st.expander(f"查看 {len(result.errors)} 条文件错误", expanded=True):
            st.dataframe(result.errors, width="stretch", hide_index=True)

    st.subheader("性能结果")
    preferred = [
        "sample_id", "group", "source_type", "rp02_mpa", "rm_mpa",
        "uniform_elongation_percent", "at_percent", "e_gpa",
        "modulus_flag", "yield_flag", "curve_note",
    ]
    columns = [name for name in preferred if any(name in row for row in result.rows)]
    table = [{name: row.get(name, "") for name in columns} for row in result.rows]
    st.dataframe(table, width="stretch", hide_index=True)
    st.caption("Rp0.2、Rm：MPa；均匀延伸率、At：%；E：GPa。空值表示自动判据未通过，请结合原始曲线核对。")

    _show_custom_plotter(result)

    images = sorted(name for name in result.files if name.lower().endswith(".png"))
    if images:
        with st.expander("查看后台自动生成的文件/分组图"):
            chosen = st.selectbox(
                "选择图片",
                images,
                format_func=lambda name: PurePosixPath(name).name,
                key="preview_image",
            )
            st.image(result.files[chosen], caption=chosen, width="stretch")

    st.subheader("下载结果")
    st.download_button(
        "下载全部结果 ZIP",
        data=result.archive,
        file_name="拉伸曲线处理结果.zip",
        mime="application/zip",
        type="primary",
        key="download_archive",
    )
    if "curve_index.csv" in result.files:
        st.download_button(
            "下载性能索引 CSV",
            data=result.files["curve_index.csv"],
            file_name="curve_index.csv",
            mime="text/csv",
            key="download_index",
        )
    with st.expander("下载单个图表或性能表"):
        downloadable = sorted(
            (name, data) for name, data in result.files.items()
            if name != "curve_index.csv" and not name.endswith("processing_report.json")
        )
        for index, (name, data) in enumerate(downloadable):
            suffix = PurePosixPath(name).suffix.lower()
            mime = {".png": "image/png", ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}.get(suffix, "application/octet-stream")
            st.download_button(
                f"下载 {name}", data=data, file_name=PurePosixPath(name).name,
                mime=mime, key=f"download_file_{index}",
            )


def _pandat_melting_text(result: PandatResult) -> str:
    if result.melting_start_C is not None and result.melting_end_C is not None:
        return f"{result.melting_start_C:.0f}–{result.melting_end_C:.0f} ℃"
    if result.melting_start_C is not None:
        return f"{result.melting_start_C:.0f} ℃起，未完全液化"
    return "未检测到液相"


def _clear_pandat_result() -> None:
    st.session_state.pandat_result = None
    st.session_state.pandat_files = None
    st.session_state.pandat_source = None
    st.session_state.pandat_processed_name = None


def _pandat_fingerprint(uploaded, sample_name: str, threshold: float) -> str:
    digest = hashlib.sha256()
    digest.update(sample_name.encode("utf-8", errors="replace"))
    digest.update(repr(threshold).encode("ascii"))
    if uploaded is None:
        digest.update(b"no-upload")
    else:
        data = uploaded.getvalue()
        digest.update(uploaded.name.encode("utf-8", errors="replace"))
        digest.update(len(data).to_bytes(8, "big"))
        digest.update(hashlib.sha256(data).digest())
    return digest.hexdigest()


def _run_pandat(data: bytes, *, filename: str, sample_name: str, threshold: float, source: str) -> None:
    try:
        with st.spinner("正在解析原生相分数、识别相区并生成图片…"):
            result = parse_pandat_line(
                data,
                filename=filename,
                phase_threshold=threshold,
            )
            files = build_download_bundle(result, sample_name=sample_name)
        st.session_state.pandat_result = result
        st.session_state.pandat_files = files
        st.session_state.pandat_source = source
        st.session_state.pandat_processed_name = sample_name
    except PandatError as exc:
        _clear_pandat_result()
        st.error(str(exc))


def _show_pandat_page() -> None:
    st.markdown(
        """
        <section class="hero">
          <div class="eyebrow">PANDAT LINE RESULT WORKBENCH</div>
          <h1>Pandat 相图在线整理</h1>
          <p>上传 Pandat 线计算导出的 DAT 文件，自动生成独立相分数曲线、25 ℃相内成分说明、相区带和熔化区间。</p>
        </section>
        """,
        unsafe_allow_html=True,
    )
    st.warning(
        "网页不运行 Pandat，也不包含热力学数据库或许可证；请先在有授权的本机完成平衡线计算，再上传导出的结果。"
    )
    if "pandat_result" not in st.session_state:
        st.session_state.pandat_result = None
    if "pandat_files" not in st.session_state:
        st.session_state.pandat_files = None
    if "pandat_source" not in st.session_state:
        st.session_state.pandat_source = None
    if "pandat_processed_name" not in st.session_state:
        st.session_state.pandat_processed_name = None
    if "pandat_input_fingerprint" not in st.session_state:
        st.session_state.pandat_input_fingerprint = ""

    st.markdown("<p class='step'>01 · 上传 Pandat 线计算结果</p>", unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "上传 Pandat DAT/TXT",
        type=["dat", "txt"],
        accept_multiple_files=False,
        max_upload_size=20,
        help="需要包含 T、f(@相名) 列；若包含 x(元素@相名) 列，网页会同时显示25 ℃相内成分。",
        key="pandat_upload",
    )
    sample_name = st.text_input(
        "样品名称",
        value="C63(C9)",
        max_chars=80,
        key="pandat_sample_name",
    )
    threshold = st.number_input(
        "相存在阈值（摩尔分数）",
        min_value=1e-12,
        max_value=1e-2,
        value=1e-8,
        format="%.1e",
        help="低于该值的数值仍保留在曲线和CSV中，但不计入相区组合。",
        key="pandat_phase_threshold",
    )
    current_fingerprint = _pandat_fingerprint(uploaded, sample_name, float(threshold))
    if current_fingerprint != st.session_state.pandat_input_fingerprint:
        if st.session_state.pandat_result is not None:
            _clear_pandat_result()
        st.session_state.pandat_input_fingerprint = current_fingerprint
    left, middle, right = st.columns([1.15, 1, 1])
    with left:
        if st.button(
            "开始生成相图",
            type="primary",
            disabled=uploaded is None,
            key="run_pandat_upload",
            width="stretch",
        ):
            _run_pandat(
                uploaded.getvalue(),
                filename=uploaded.name,
                sample_name=sample_name.strip() or PurePosixPath(uploaded.name).stem,
                threshold=float(threshold),
                source="upload",
            )
    with middle:
        if st.button("运行 C63(C9) 示例", key="run_pandat_demo", width="stretch"):
            demo = Path(__file__).with_name("examples") / "C63_C9_line_up.dat"
            _run_pandat(
                demo.read_bytes(),
                filename=demo.name,
                sample_name=sample_name.strip() or "C63(C9)",
                threshold=float(threshold),
                source="demo",
            )
    with right:
        if st.button(
            "清除当前结果",
            disabled=st.session_state.pandat_result is None,
            key="clear_pandat_result",
            width="stretch",
        ):
            _clear_pandat_result()

    with st.expander("Pandat 导出格式说明"):
        st.markdown(
            """
            - 在 Pandat 中完成 25–1500 ℃或其它温区的线计算后，导出制表符结果文件。
            - 文件必须包含 `T` 和至少一个 `f(@相名)` 列；`T` 单位行写 `K` 或 `C/℃`。
            - `x(元素@相名)` 列用于生成室温相内成分说明；缺少这些列仍可绘制相分数曲线。
            - 相变温度处的重复行会按最后一行归并，避免生成零宽度伪相区。
            - 网页展示的是热力学平衡预测，不等同于实际冷却组织或实验验证。
            """
        )

    st.markdown("<p class='step'>02 · 查看并下载</p>", unsafe_allow_html=True)
    result = st.session_state.pandat_result
    files = st.session_state.pandat_files
    if result is None or files is None:
        st.info("上传 Pandat 线计算结果后点击“开始生成相图”，或运行 C63(C9) 示例查看效果。")
        return
    st.success(f"处理完成：识别到 {len(result.phases)} 个相、{len(result.regions)} 个原生相区。")
    if st.session_state.pandat_source == "demo":
        st.info("当前为已验证的 C63(C9) 示例数据，熔化区间约为 1179–1372 ℃。")
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("温度范围", f"{result.temperatures_C[0]:.0f}–{result.temperatures_C[-1]:.0f} ℃")
    col2.metric("相数", str(len(result.phases)))
    col3.metric("相区数", str(len(result.regions)))
    col4.metric("熔化区间", _pandat_melting_text(result))
    processed_name = st.session_state.pandat_processed_name or result.filename
    st.image(files["phase_diagram.png"], caption=f"{processed_name} · Pandat 平衡相分数", width="stretch")

    st.subheader("相区表")
    st.dataframe(
        [
            {
                "起始温度 (℃)": round(region.lower_C, 6),
                "终止温度 (℃)": round(region.upper_C, 6),
                "原生相组合": " + ".join(region.phases),
            }
            for region in result.regions
        ],
        width="stretch",
        hide_index=True,
    )
    st.subheader("下载结果")
    download_columns = st.columns(5)
    downloads = [
        ("下载 PNG", "phase_diagram.png", "image/png"),
        ("下载 PDF", "phase_diagram.pdf", "application/pdf"),
        ("下载相分数 CSV", "phase_fractions.csv", "text/csv"),
        ("下载相区 CSV", "phase_regions.csv", "text/csv"),
        ("下载元数据 JSON", "phase_metadata.json", "application/json"),
    ]
    for index, (label, filename, mime) in enumerate(downloads):
        with download_columns[index]:
            st.download_button(
                label,
                data=files[filename],
                file_name=filename,
                mime=mime,
                key=f"download_pandat_{index}",
                width="stretch",
            )
    st.caption("上传数据仅用于当前会话内存处理；导出结果不会包含原始 DAT 文件。")


tool = st.radio(
    "选择工具",
    ["拉伸曲线", "Pandat 相图"],
    horizontal=True,
    key="tool_selector",
)
if tool == "Pandat 相图":
    _show_pandat_page()
    st.stop()


st.markdown(
    """
    <section class="hero">
      <div class="eyebrow">TENSILE CURVE WORKBENCH</div>
      <h1>拉伸曲线在线处理</h1>
      <p>上传仪器导出的 Excel、CSV 或保留原目录结构的 ZIP，自动完成曲线清理和性能计算，并自由组合生成论文风格图片。</p>
    </section>
    """,
    unsafe_allow_html=True,
)

if "result" not in st.session_state:
    st.session_state.result = None
if "result_source" not in st.session_state:
    st.session_state.result_source = None
if "upload_fingerprint" not in st.session_state:
    st.session_state.upload_fingerprint = ""

st.markdown("<p class='step'>01 · 选择数据</p>", unsafe_allow_html=True)
uploaded_files = st.file_uploader(
    "上传 XLSX、CSV 或 ZIP",
    type=["xlsx", "csv", "zip"],
    accept_multiple_files=True,
    max_upload_size=20,
    help="每个文件不超过 20 MiB，总计不超过 60 MiB；ZIP 可保留批次目录结构。",
    key="input_files",
)
current_fingerprint = _fingerprint(uploaded_files)
if current_fingerprint != st.session_state.upload_fingerprint:
    if st.session_state.result_source == "upload":
        _clear_result()
    st.session_state.upload_fingerprint = current_fingerprint

if uploaded_files:
    st.markdown(
        f"<div class='notice'>已选择 {len(uploaded_files)} 个上传项。点击“开始处理”后，文件将在独立临时目录中运行，完成后自动清理。</div>",
        unsafe_allow_html=True,
    )

left, middle, right = st.columns([1.15, 1, 1])
with left:
    if st.button("开始处理", type="primary", disabled=not uploaded_files, key="run_uploads", width="stretch"):
        _run(_uploaded_bytes(uploaded_files), "upload")
with middle:
    if st.button("运行合成示例", key="run_demo", width="stretch"):
        _run(make_demo_uploads(), "demo")
with right:
    clear_disabled = st.session_state.result is None
    if st.button("清除当前结果", disabled=clear_disabled, key="clear_results", width="stretch"):
        _clear_result()

with st.expander("输入格式与结果说明"):
    st.markdown(
        """
        - Excel 曲线区：第 1 行为曲线名，第 2 行为相邻的应变/应力单位，第 3 行起为数值；支持 `%`、`mm/mm` 和 `MPa`、`kN/mm²`。
        - CSV 需包含中文曲线表头和 `%`、`MPa` 单位行。散装 CSV 会自动放入临时 `_Exports` 目录。
        - 多批次或依赖同文件夹 `Test Run` 汇总匹配时，请上传 ZIP 保留目录层级。
        - 自动 E、Rp0.2 与断裂模式是辅助计算；结果需结合原始曲线、试验条件和仪器结果核对。
        - 同一试验若同时出现在 Excel 与 CSV 中，当前版本不会自动去重。
        """
    )

st.markdown("<p class='step'>02 · 查看并下载</p>", unsafe_allow_html=True)
if st.session_state.result is None:
    st.info("上传数据后点击“开始处理”，或先运行合成示例查看页面效果。")
else:
    if st.session_state.result_source == "demo":
        st.info("当前显示的是合成演示数据，仅验证处理流程，不代表真实实验精度。")
    _show_results(st.session_state.result)

st.divider()
st.caption("上传数据只用于本次会话的临时处理；下载包不包含原始上传文件。")
