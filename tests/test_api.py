"""
Tests for main.py -- deliberately avoid live Nominatim calls (CI runners
are cloud IPs, which Nominatim's public server may block per its usage
policy -- see README.md's "Nominatim IP-blocking" section). Instead this
mocks geocoder_core.geocode_address_multi_tier so we're testing our own
auth/validation/caching/error-shaping logic, not network reachability.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

os.environ["GEOCODE_API_KEY"] = "test-key"
os.environ["NOMINATIM_CONTACT_EMAIL"] = "test@example.com"
os.environ["CACHE_DB_PATH"] = ":memory:"  # avoid leaving cache files around between test runs

import pytest
from fastapi.testclient import TestClient

import main as api_module

client = TestClient(api_module.app)


def _fake_geocode_result(**overrides):
    base = {
        "Latitude": 29.7604, "Longitude": -95.3698, "Confidence_Level": "EXACT_STREET",
        "Standardized_Address": "123 Main St, Houston, TX, 77002",
        "Match_Method": "structured_candidate_0", "Comment": "Matched at exact street-address level.",
        "Quality_Flags": [],
    }
    base.update(overrides)
    return base


def test_health_check_no_auth_required():
    resp = client.get("/health")
    assert resp.status_code == 200
    assert resp.json() == {"status": "ok"}


def test_geocode_missing_key_returns_422():
    resp = client.get("/geocode", params={"address": "123 Main St"})
    assert resp.status_code == 422  # FastAPI's own required-query-param validation


def test_geocode_wrong_key_returns_401():
    resp = client.get("/geocode", params={"address": "123 Main St", "key": "wrong-key"})
    assert resp.status_code == 401
    assert resp.json() == {"error": "Invalid or missing API key"}


def test_geocode_empty_address_returns_400(monkeypatch):
    resp = client.get("/geocode", params={"address": "   ", "key": "test-key"})
    assert resp.status_code in (400, 422)


def test_geocode_success_shape(monkeypatch):
    monkeypatch.setattr(
        api_module.gc, "geocode_address_multi_tier",
        lambda address, **kwargs: _fake_geocode_result(),
    )
    resp = client.get("/geocode", params={"address": "123 Main St, Houston, TX, 77002", "key": "test-key"})
    assert resp.status_code == 200
    data = resp.json()
    assert data["lat"] == 29.7604
    assert data["lon"] == -95.3698
    assert data["confidence"] == "EXACT_STREET"
    assert data["cached"] is False


def test_geocode_second_call_is_cached(monkeypatch):
    call_count = {"n": 0}

    def _counting_geocode(address, **kwargs):
        call_count["n"] += 1
        return _fake_geocode_result()

    monkeypatch.setattr(api_module.gc, "geocode_address_multi_tier", _counting_geocode)

    addr = "999 Cache Test Ave, Austin, TX, 78701"
    r1 = client.get("/geocode", params={"address": addr, "key": "test-key"})
    r2 = client.get("/geocode", params={"address": addr, "key": "test-key"})

    assert r1.json()["cached"] is False
    assert r2.json()["cached"] is True
    assert call_count["n"] == 1  # second call must NOT hit the geocoder again


def test_batch_geocode_respects_max_100(monkeypatch):
    monkeypatch.setattr(
        api_module.gc, "geocode_address_multi_tier",
        lambda address, **kwargs: _fake_geocode_result(),
    )
    addresses = "|".join(f"{i} Test St" for i in range(101))
    resp = client.get("/batch_geocode", params={"addresses": addresses, "key": "test-key"})
    assert resp.status_code == 400


def test_batch_geocode_success(monkeypatch):
    monkeypatch.setattr(
        api_module.gc, "geocode_address_multi_tier",
        lambda address, **kwargs: _fake_geocode_result(),
    )
    resp = client.get(
        "/batch_geocode",
        params={"addresses": "1 First St|2 Second St", "key": "test-key"},
    )
    assert resp.status_code == 200
    results = resp.json()["results"]
    assert len(results) == 2
    assert all(r["confidence"] == "EXACT_STREET" for r in results)


def test_network_error_returns_502(monkeypatch):
    import requests as requests_module

    def _raise_network_error(address, **kwargs):
        raise requests_module.ConnectionError("simulated network failure")

    monkeypatch.setattr(api_module.gc, "geocode_address_multi_tier", _raise_network_error)
    resp = client.get(
        "/geocode",
        params={"address": "999 Network Fail Test Rd, Nowhere, TX, 00000", "key": "test-key"},
    )
    assert resp.status_code == 502
    assert "error" in resp.json()
