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
