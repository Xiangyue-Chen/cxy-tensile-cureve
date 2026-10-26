"""Streamlit front end for the isolated tensile-curve processor."""
from __future__ import annotations

import hashlib
from pathlib import PurePosixPath

import streamlit as st

from tensile_web import ProcessingError, ProcessingResult, Upload, make_demo_uploads, process_uploads


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


def _run(uploads: list[Upload], source: str) -> None:
    try:
        with st.spinner("正在识别曲线、计算性能并生成结果…"):
            st.session_state.result = process_uploads(uploads)
        st.session_state.result_source = source
    except ProcessingError as exc:
        _clear_result()
        st.error(str(exc))


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

    images = sorted(name for name in result.files if name.lower().endswith(".png"))
    if images:
        st.subheader("曲线预览")
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


st.markdown(
    """
    <section class="hero">
      <div class="eyebrow">TENSILE CURVE WORKBENCH</div>
      <h1>拉伸曲线在线处理</h1>
      <p>上传仪器导出的 Excel、CSV 或保留原目录结构的 ZIP，自动完成曲线清理、性能计算、分组绘图和结果打包。</p>
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
