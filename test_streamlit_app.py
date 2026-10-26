from pathlib import Path

from streamlit.testing.v1 import AppTest


APP = Path(__file__).with_name("streamlit_app.py")


def test_empty_page_requires_files_before_processing():
    app = AppTest.from_file(str(APP)).run()
    assert not app.exception
    assert app.button(key="run_uploads").disabled
    assert not app.get("download_button")
    assert not app.metric


def test_demo_runs_real_pipeline_and_clear_removes_results():
    app = AppTest.from_file(str(APP)).run()
    app.button(key="run_demo").click().run(timeout=60)
    assert not app.exception
    assert app.success
    assert len(app.dataframe) >= 1
    assert int(app.metric[0].value) > 0
    assert len(app.multiselect(key="plot_curve_selection").value) == 2
    assert app.radio(key="plot_mode").value == "叠加对比"
    assert app.checkbox(key="mark_performance_points").value is True
    assert app.text_input(key="plot_title").value == "Tensile stress–strain curves"
    assert len(app.get("download_button")) >= 3
    app.button(key="clear_results").click().run()
    assert not app.exception
    assert not app.get("download_button")
    assert not app.metric


def test_changed_upload_fingerprint_invalidates_old_upload_result():
    app = AppTest.from_file(str(APP)).run()
    app.button(key="run_demo").click().run(timeout=60)
    assert app.metric
    app.session_state["result_source"] = "upload"
    app.session_state["upload_fingerprint"] = "previous-upload"
    app.run()
    assert not app.exception
    assert not app.metric
    assert not app.get("download_button")


def test_pandat_page_runs_c63_demo_and_offers_plot_downloads():
    app = AppTest.from_file(str(APP)).run()

    app.radio(key="tool_selector").set_value("Pandat 相图").run()

    assert not app.exception
    assert any("Pandat 线计算" in element.value for element in app.markdown)
    assert app.button(key="run_pandat_demo")
    app.button(key="run_pandat_demo").click().run(timeout=60)
    assert not app.exception
    assert app.success
    metrics = {metric.label: metric.value for metric in app.metric}
    assert metrics["相数"] == "8"
    assert metrics["相区数"] == "12"
    assert "1179" in metrics["熔化区间"]
    assert "1372" in metrics["熔化区间"]
    assert len(app.dataframe) >= 1
    download_labels = {button.label for button in app.get("download_button")}
    assert {"下载 PNG", "下载 PDF", "下载相分数 CSV", "下载相区 CSV", "下载元数据 JSON"} <= download_labels


def test_changed_pandat_inputs_invalidate_previous_result():
    app = AppTest.from_file(str(APP)).run()
    app.radio(key="tool_selector").set_value("Pandat 相图").run()
    app.button(key="run_pandat_demo").click().run(timeout=60)
    assert app.metric
    assert len(app.get("download_button")) == 5

    app.text_input(key="pandat_sample_name").set_value("Different-Sample").run()

    assert not app.exception
    assert not app.metric
    assert not app.get("download_button")
