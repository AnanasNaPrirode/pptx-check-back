import json
import logging
import time
import zipfile
from io import BytesIO

import pytest
from fastapi.testclient import TestClient

from app.config import TIMEOUT_MESSAGE
from app.main import create_app
from tests.pptx_factory import add_text, blank, presentation, save


@pytest.fixture()
def client():
    return TestClient(create_app(), raise_server_exceptions=False)


def _post(client, data: bytes, filename="talk.pptx", content_type=None, checks=None, headers=None):
    files = {
        "file": (
            filename,
            data,
            content_type or "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
    }
    form = {}
    if checks is not None:
        form["enabled_checks"] = json.dumps(checks)
    return client.post("/api/check", files=files, data=form, headers=headers)


def test_checks_catalog(client):
    response = client.get("/api/checks")
    assert response.status_code == 200
    payload = response.json()
    assert [item["id"] for item in payload] == [
        "aspect_16_9",
        "embedded_fonts",
        "slide_duplicates",
        "missing_title",
        "title_dot",
        "has_video",
        "quote_guillemets",
        "em_dash",
        "ordinals",
        "list_periods",
        "title_heuristic",
        "font_size_min",
        "bottom_margin",
        "number_grouping",
        "english_ratio",
        "superscript_units",
        "overlap",
        "contrast_fill",
        "hanging_prepositions",
    ]
    heuristic = next(item for item in payload if item["id"] == "title_heuristic")
    assert heuristic["default_enabled"] is False
    assert heuristic["level"] == "warning"
    assert payload[0]["level"] == "error"
    assert payload[3]["level"] == "error"


def test_successful_report_shape(client):
    prs = presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Тема доклада"
    response = _post(client, save(prs))
    assert response.status_code == 200, response.text
    body = response.json()
    assert body["filename"] == "talk.pptx"
    assert body["slides_count"] == 1
    assert "title_heuristic" not in body["applied_checks"]
    assert "font_size_min" in body["applied_checks"]
    slide_payload = body["slides"][0]
    assert slide_payload["width_emu"] == 12_192_000
    assert slide_payload["number"] == 1
    assert any(element.get("is_title") for element in slide_payload["elements"])


def test_disabled_check_is_not_applied(client):
    prs = presentation()
    prs.slide_width = 9_144_000
    slide = blank(prs)
    add_text(slide, "Мелкий", 400_000, 400_000, 4_000_000, 800_000, size=14)
    response = _post(client, save(prs), checks=["font_size_min"])
    body = response.json()
    assert response.status_code == 200
    assert body["applied_checks"] == ["font_size_min"]
    assert body["presentation_issues"] == []
    assert body["slides"][0]["issues"][0]["check_id"] == "font_size_min"


def test_unknown_check(client):
    prs = presentation()
    blank(prs)
    response = _post(client, save(prs), checks=["font_size_min", "not_a_check"])
    assert response.status_code == 422
    assert response.json()["error"]["code"] == "UNKNOWN_CHECK"


def test_unsupported_extension(client):
    response = _post(client, b"PK\x03\x04", filename="notes.pdf")
    assert response.status_code == 415
    assert response.json()["error"]["code"] == "UNSUPPORTED_TYPE"


def test_unsupported_mime(client):
    prs = presentation()
    blank(prs)
    response = _post(client, save(prs), content_type="image/png")
    assert response.status_code == 415
    assert response.json()["error"]["code"] == "UNSUPPORTED_TYPE"


def test_invalid_file(client):
    response = _post(client, b"this is not a pptx")
    assert response.status_code == 400
    assert response.json()["error"]["code"] == "INVALID_FILE"


def test_invalid_ooxml(client):
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("readme.txt", "hello")
    response = _post(client, buffer.getvalue())
    assert response.status_code == 400
    assert response.json()["error"]["code"] == "INVALID_FILE"


def test_macros_rejected(client):
    prs = presentation()
    blank(prs)
    buffer = BytesIO(save(prs))
    with zipfile.ZipFile(buffer, "a") as archive:
        archive.writestr("ppt/vbaProject.bin", b"macro")
    response = _post(client, buffer.getvalue())
    assert response.status_code == 400
    assert response.json()["error"]["code"] == "MACROS_NOT_ALLOWED"


def test_zip_slip_rejected(client):
    prs = presentation()
    blank(prs)
    buffer = BytesIO(save(prs))
    with zipfile.ZipFile(buffer, "a") as archive:
        archive.writestr("../evil.txt", b"x")
    response = _post(client, buffer.getvalue())
    assert response.status_code == 400
    assert response.json()["error"]["code"] == "INVALID_FILE"


def test_archive_too_large(client, monkeypatch):
    monkeypatch.setattr("app.config.MAX_UNCOMPRESSED_BYTES", 100)
    prs = presentation()
    blank(prs)
    response = _post(client, save(prs))
    assert response.status_code == 400
    assert response.json()["error"]["code"] == "ARCHIVE_TOO_LARGE"


def test_file_too_large(client, monkeypatch):
    monkeypatch.setattr("app.config.MAX_FILE_BYTES", 50)
    prs = presentation()
    blank(prs)
    response = _post(client, save(prs))
    assert response.status_code == 413
    assert response.json()["error"]["code"] == "FILE_TOO_LARGE"


def test_content_length_too_large(client, monkeypatch):
    monkeypatch.setattr("app.config.MAX_CONTENT_LENGTH_BYTES", 10)
    prs = presentation()
    blank(prs)
    response = _post(client, save(prs))
    assert response.status_code == 413
    assert response.json()["error"]["code"] == "FILE_TOO_LARGE"


def test_rate_limit(client):
    prs = presentation()
    blank(prs)
    data = save(prs)
    headers = {"X-Forwarded-For": "203.0.113.10"}
    first = _post(client, data, checks=[], headers=headers)
    second = _post(client, data, checks=[], headers=headers)
    assert first.status_code == 200
    assert second.status_code == 429
    assert second.json()["error"]["code"] == "RATE_LIMITED"
    assert int(second.headers["Retry-After"]) >= 1


def test_timeout(client, monkeypatch):
    def slow(*_args, **_kwargs):
        time.sleep(1)

    monkeypatch.setattr("app.main.build_report", slow)
    monkeypatch.setattr("app.config.CHECK_TIMEOUT_SECONDS", 0.2)
    prs = presentation()
    blank(prs)
    response = _post(client, save(prs), checks=[])
    assert response.status_code == 504
    assert response.json()["error"] == {"code": "TIMEOUT", "message": TIMEOUT_MESSAGE}


def test_log_has_no_filename(client, caplog):
    prs = presentation()
    blank(prs)
    with caplog.at_level(logging.INFO, logger="pptx_checker"):
        response = _post(client, save(prs), filename="secret-talk.pptx", checks=[])
    assert response.status_code == 200
    assert "secret-talk" not in caplog.text


def test_internal_error_hides_details(client, monkeypatch):
    def boom(*_args, **_kwargs):
        raise RuntimeError("secret-talk-body")

    monkeypatch.setattr("app.main.build_report", boom)
    prs = presentation()
    blank(prs)
    response = _post(client, save(prs), checks=[])
    assert response.status_code == 500
    assert response.json()["error"]["code"] == "INTERNAL_ERROR"
    assert "secret-talk-body" not in response.text


def test_check_error_does_not_fail_request(client, monkeypatch):
    real = __import__("app.pipeline", fromlist=["run_check"]).run_check

    def wrapped(check, model):
        if check.id == "aspect_16_9":
            raise RuntimeError("boom")
        real(check, model)

    monkeypatch.setattr("app.pipeline.run_check", wrapped)
    prs = presentation()
    prs.slide_width = 9_144_000
    slide = blank(prs)
    add_text(slide, "Мелкий текст", 400_000, 400_000, 4_000_000, 800_000, size=14)
    response = _post(client, save(prs), checks=["aspect_16_9", "font_size_min"])
    assert response.status_code == 200
    body = response.json()
    assert body["presentation_issues"] == []
    assert body["slides"][0]["issues"][0]["check_id"] == "font_size_min"
