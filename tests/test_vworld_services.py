"""Synthetic contract tests; these fixtures are NOT captured VWorld responses."""

import io
import json
import os
import threading
import urllib.error
from datetime import date
from unittest.mock import Mock
from urllib.parse import parse_qs, urlsplit

import pytest

from services.api_key_service import ApiKeyError, ApiKeyService
from services.vworld_geocoding_service import GeocodingError, VWorldGeocodingService


def payload(**fields):
    return {"response": {"status": "OK", "result": [{"type": "parcel", "text": "주소",
            "structure": {"level1": "서울특별시", **fields}}]}}


def test_request_coordinate_order_and_code_preservation():
    transport = Mock(return_value=io.BytesIO(json.dumps(payload(level4LC="0012345678")).encode()))
    result = VWorldGeocodingService(transport).lookup("37.5", "127", "test-secret")
    query = parse_qs(urlsplit(transport.call_args.args[0].full_url).query)
    assert query == {k: [v] for k, v in {
        "service": "address", "request": "getAddress", "version": "2.0", "crs": "EPSG:4326",
        "point": "127,37.5", "type": "PARCEL", "format": "json", "simple": "false", "key": "test-secret",
    }.items()}
    assert result.structure["level4LC"] == "0012345678"
    assert result.structure["level5"] == ""


@pytest.mark.parametrize("lat,lon", [(None, 127), ("", 127), (91, 127), (37, 181), ("nan", 127), (37, "inf")])
def test_invalid_coordinates_do_not_request(lat, lon):
    transport = Mock()
    with pytest.raises(GeocodingError) as error:
        VWorldGeocodingService(transport).lookup(lat, lon, "key")
    assert error.value.code == "coordinates"
    transport.assert_not_called()


def test_missing_key_does_not_request():
    transport = Mock()
    with pytest.raises(GeocodingError) as error:
        VWorldGeocodingService(transport).lookup(37, 127, "")
    assert error.value.code == "missing_key"
    transport.assert_not_called()


@pytest.mark.parametrize("body,code", [
    ({"response": {"status": "NOT_FOUND"}}, "not_found"),
    ({"response": {"status": "ERROR", "error": {"code": "INVALID_KEY", "text": "test-secret"}}}, "authentication"),
    ({"response": {"status": "ERROR", "error": {"code": "OVER_QUERY_LIMIT"}}}, "rate_limit"),
    ({"response": {"status": "ERROR"}}, "api_error"),
    ({}, "invalid_response"),
    ({"response": {"status": "OK", "result": []}}, "invalid_response"),
    ({"response": {"status": "OK", "result": [{"structure": {}}]}}, "invalid_response"),
    (payload(level4LC=1234), "invalid_response"),
])
def test_response_failures(body, code):
    with pytest.raises(GeocodingError) as error:
        VWorldGeocodingService.parse_response(body)
    assert error.value.code == code
    assert "test-secret" not in str(error.value)


@pytest.mark.parametrize("failure,code", [
    (TimeoutError("test-secret"), "timeout"),
    (urllib.error.URLError(TimeoutError()), "timeout"),
    (urllib.error.URLError("test-secret"), "network"),
    (urllib.error.HTTPError("https://example.test/?key=test-secret", 403, "test-secret", {}, None), "authentication"),
    (urllib.error.HTTPError("https://example.test/?key=test-secret", 429, "test-secret", {}, None), "rate_limit"),
    (urllib.error.HTTPError("https://example.test/?key=test-secret", 503, "test-secret", {}, None), "http"),
], ids=["timeout", "url-timeout", "network", "auth", "limit", "server"])
def test_transport_errors_are_sanitized(failure, code):
    with pytest.raises(GeocodingError) as error:
        VWorldGeocodingService(Mock(side_effect=failure)).lookup(37, 127, "test-secret")
    assert error.value.code == code
    assert "test-secret" not in str(error.value)


def test_invalid_json():
    with pytest.raises(GeocodingError) as error:
        VWorldGeocodingService(Mock(return_value=io.BytesIO(b"<html>"))).lookup(37, 127, "key")
    assert error.value.code == "invalid_response"


class MemoryVault:
    def __init__(self):
        self.data = {}

    def get_password(self, service, account):
        return self.data.get((service, account))

    def set_password(self, service, account, value):
        self.data[service, account] = value

    def delete_password(self, service, account):
        del self.data[service, account]


def test_key_lifecycle_and_restart():
    vault = MemoryVault()
    keys = ApiKeyService(vault)
    assert keys.load() == ("", None)
    keys.save("first", date(2027, 1, 1))
    assert ApiKeyService(vault).load() == ("first", date(2027, 1, 1))
    keys.save("second")
    assert keys.load() == ("second", None)
    assert list(vault.data) == [("BioFlow.VWorld.Address", "api-key")]
    keys.delete()
    keys.delete()
    assert keys.load() == ("", None)


@pytest.mark.parametrize("operation", ["load", "save", "delete"])
def test_vault_failures_are_sanitized(operation):
    vault = Mock()
    vault.get_password.side_effect = RuntimeError("test-secret")
    vault.set_password.side_effect = RuntimeError("test-secret")
    with pytest.raises(ApiKeyError) as error:
        getattr(ApiKeyService(vault), operation)(*( ["test-secret"] if operation == "save" else []))
    assert "test-secret" not in str(error.value)


def test_settings_connection_is_nonblocking_and_does_not_save_key():
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    from PySide6.QtWidgets import QApplication
    from ui.dialogs.api_settings_dialog import ApiSettingsDialog
    app = QApplication.instance() or QApplication([])
    entered = threading.Event()
    release = threading.Event()
    thread_ids = []

    def lookup(*args):
        thread_ids.append(threading.get_ident())
        entered.set()
        release.wait(3)

    vault = MemoryVault()
    dialog = ApiSettingsDialog(key_service=ApiKeyService(vault), geocoder=Mock(lookup=lookup))
    dialog.key_input.setText("test-secret")
    try:
        dialog.test_connection()
        assert entered.wait(1)
        assert not dialog.test_button.isEnabled()
        assert thread_ids[0] != threading.get_ident()
        assert vault.data == {}
        app.processEvents()
    finally:
        release.set()
        assert dialog.worker.wait(3000)
        app.processEvents()
        dialog.reject()
    assert dialog.key_input.text() == ""
    assert dialog.worker.key == ""
