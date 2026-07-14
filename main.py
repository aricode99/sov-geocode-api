"""
main.py -- FastAPI wrapper around geocoder_core.py

Exposes the existing cascading Nominatim geocoder as a small HTTP API,
so it can be called directly from Excel (WEBSERVICE, Power Query, or a
365 LAMBDA custom function) with zero client-side installation.

geocoder_core.py itself is untouched -- all API-specific concerns
(auth, caching, rate limiting, HTTP error shaping) live here.
"""

import hashlib
import json
import os
import sqlite3
import threading
import time
from contextlib import asynccontextmanager

import requests
from fastapi import FastAPI, HTTPException, Query
from fastapi.responses import JSONResponse

import geocoder_core as gc

# ─────────────────────────────────────────────────────────────────────────
# Config (env vars so nothing sensitive is hardcoded / committed to git)
# ─────────────────────────────────────────────────────────────────────────
API_KEY = os.environ.get("GEOCODE_API_KEY", "")
CONTACT_EMAIL = os.environ.get("NOMINATIM_CONTACT_EMAIL", "not-configured@example.com")
MIN_REQUEST_INTERVAL = float(os.environ.get("MIN_REQUEST_INTERVAL_SECS", "1.1"))
CACHE_TTL_SECONDS = int(os.environ.get("CACHE_TTL_SECONDS", str(60 * 60 * 24 * 90)))  # 90 days
CACHE_DB_PATH = os.environ.get("CACHE_DB_PATH", "geocode_cache.sqlite3")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USER_AGENT = f"SOV-Geocode-API/1.0 (contact: {CONTACT_EMAIL})"


# ─────────────────────────────────────────────────────────────────────────
# SQLite cache (see architecture doc section G -- same design)
# ─────────────────────────────────────────────────────────────────────────
_cache_lock = threading.Lock()


def _init_cache():
    conn = sqlite3.connect(CACHE_DB_PATH, check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS geocode_cache (
            address_hash TEXT PRIMARY KEY,
            raw_address TEXT,
            result_json TEXT,
            cached_at REAL
        )
        """
    )
    conn.commit()
    return conn


_cache_conn = _init_cache()


def _hash_address(address: str) -> str:
    normalized = " ".join(address.strip().lower().split())
    return hashlib.sha256(normalized.encode("utf-8")).hexdigest()


def cache_lookup(address: str):
    h = _hash_address(address)
    with _cache_lock:
        row = _cache_conn.execute(
            "SELECT result_json, cached_at FROM geocode_cache WHERE address_hash = ?", (h,)
        ).fetchone()
    if not row:
        return None
    result_json, cached_at = row
    if time.time() - cached_at > CACHE_TTL_SECONDS:
        return None
    return json.loads(result_json)


def cache_store(address: str, result: dict):
    h = _hash_address(address)
    with _cache_lock:
        _cache_conn.execute(
            "INSERT OR REPLACE INTO geocode_cache (address_hash, raw_address, result_json, cached_at) "
            "VALUES (?, ?, ?, ?)",
            (h, address, json.dumps(result), time.time()),
        )
        _cache_conn.commit()


# ─────────────────────────────────────────────────────────────────────────
# Rate limiter -- shared across ALL requests to this API instance, so it
# collectively respects Nominatim's ~1 req/sec policy regardless of how
# many Excel users/formulas are calling in simultaneously.
# ─────────────────────────────────────────────────────────────────────────
class RateLimiter:
    def __init__(self, min_interval_secs: float):
        self._min_interval = min_interval_secs
        self._lock = threading.Lock()
        self._last_call = 0.0

    def wait(self):
        with self._lock:
            elapsed = time.monotonic() - self._last_call
            if elapsed < self._min_interval:
                time.sleep(self._min_interval - elapsed)
            self._last_call = time.monotonic()


_rate_limiter = RateLimiter(MIN_REQUEST_INTERVAL)


# ─────────────────────────────────────────────────────────────────────────
# App lifespan -- load the ZIP reference table once at startup
# ─────────────────────────────────────────────────────────────────────────
@asynccontextmanager
async def lifespan(app: FastAPI):
    zip_path = os.path.join(BASE_DIR, "reference_data", "USZIP.xlsx")
    summary = gc.load_zip_reference(zip_path)
    print(f"Loaded ZIP reference: {summary}")
    yield


app = FastAPI(
    title="SOV Address Geocoder API",
    description="Thin HTTP wrapper around geocoder_core.py, for calling from Excel.",
    version="1.0.0",
    lifespan=lifespan,
)


def _check_api_key(key: str):
    if not API_KEY:
        # No key configured server-side -- fail closed, not open, so a
        # misconfigured deploy doesn't accidentally become a public
        # unauthenticated Nominatim proxy.
        raise HTTPException(status_code=500, detail="Server misconfigured: no API key set")
    if key != API_KEY:
        raise HTTPException(status_code=401, detail="Invalid or missing API key")


@app.get("/")
def root():
    """Friendly landing response -- avoids a confusing 404 when Render's
    own platform (or a person) hits the bare URL directly."""
    return {
        "service": "SOV Address Geocoder API",
        "endpoints": ["/health", "/geocode?address=...&key=...",
                      "/batch_geocode?addresses=...&key=..."],
    }


@app.get("/health")
def health():
    """Unauthenticated health check -- for the hosting platform's uptime probe."""
    return {"status": "ok"}


@app.get("/geocode")
def geocode(
    address: str = Query(..., min_length=1, max_length=500, description="Full address, comma-separated"),
    key: str = Query(..., description="API key"),
):
    _check_api_key(key)

    address = address.strip()
    if not address:
        raise HTTPException(status_code=400, detail="Address is empty")

    cached = cache_lookup(address)
    if cached is not None:
        return {**cached, "cached": True}

    _rate_limiter.wait()
    try:
        result = gc.geocode_address_multi_tier(address, user_agent=USER_AGENT, sleep_secs=MIN_REQUEST_INTERVAL)
    except requests.RequestException as e:
        raise HTTPException(status_code=502, detail=f"Network error contacting Nominatim: {e}")
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Unexpected geocoding error: {e}")

    payload = {
        "address": address,
        "lat": result["Latitude"],
        "lon": result["Longitude"],
        "confidence": result["Confidence_Level"],
        "standardized": result["Standardized_Address"],
        "match_method": result["Match_Method"],
        "comment": result["Comment"],
    }
    cache_store(address, payload)
    return {**payload, "cached": False}


@app.get("/batch_geocode")
def batch_geocode(
    addresses: str = Query(..., description="Pipe-separated (|) list of addresses"),
    key: str = Query(..., description="API key"),
):
    """Convenience endpoint for Power Query's List.Transform pattern, or a
    single call covering several addresses at once. Still respects the
    shared rate limiter sequentially -- this does not parallelize calls
    to Nominatim."""
    _check_api_key(key)

    addr_list = [a.strip() for a in addresses.split("|") if a.strip()]
    if not addr_list:
        raise HTTPException(status_code=400, detail="No addresses provided")
    if len(addr_list) > 100:
        raise HTTPException(status_code=400, detail="Max 100 addresses per batch call")

    results = []
    for addr in addr_list:
        cached = cache_lookup(addr)
        if cached is not None:
            results.append({**cached, "cached": True})
            continue
        _rate_limiter.wait()
        try:
            result = gc.geocode_address_multi_tier(addr, user_agent=USER_AGENT, sleep_secs=MIN_REQUEST_INTERVAL)
            payload = {
                "address": addr,
                "lat": result["Latitude"],
                "lon": result["Longitude"],
                "confidence": result["Confidence_Level"],
                "standardized": result["Standardized_Address"],
                "match_method": result["Match_Method"],
                "comment": result["Comment"],
            }
            cache_store(addr, payload)
            results.append({**payload, "cached": False})
        except Exception as e:
            results.append({"address": addr, "error": str(e)})

    return {"results": results}


@app.exception_handler(HTTPException)
async def http_exception_handler(request, exc):
    # Consistent JSON error shape -- easier for Excel/Power Query to parse
    # a predictable {"error": "..."} regardless of which check failed.
    return JSONResponse(status_code=exc.status_code, content={"error": exc.detail})


if __name__ == "__main__":
    import uvicorn

    port = int(os.environ.get("PORT", "8000"))
    uvicorn.run(app, host="0.0.0.0", port=port)
