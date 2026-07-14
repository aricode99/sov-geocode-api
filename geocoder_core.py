"""
geocoder_core.py
-----------------
Core address-cleaning, parsing, and geocoding logic for the SOV Address
Geocoder tool. Ported from geocode_addresses_v8.ipynb, unchanged in
behavior -- this is the same cascading parser + Nominatim geocoder, just
organized as an importable module instead of notebook cells so it can be
driven from the Streamlit app in app.py.

No Streamlit imports here on purpose -- this module stays UI-agnostic and
testable on its own.
"""

import re
import time
from typing import Optional

import pandas as pd
import requests

try:
    import usaddress
    USADDRESS_AVAILABLE = True
except ImportError:
    USADDRESS_AVAILABLE = False


# ─────────────────────────────────────────────────────────────────────────
# Config (defaults -- app.py may override NOMINATIM/contact settings)
# ─────────────────────────────────────────────────────────────────────────
NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
DEFAULT_USER_AGENT = "SOV-Address-Geocoder/1.0 (contact: set-your-email@example.com)"
SLEEP_SECS = 1.1
COUNTRY_FILTER = "us"

LOW_CONFIDENCE = {"NO_MATCH", "STATE_LEVEL", "COUNTY_LEVEL"}

OUTPUT_COLUMNS = [
    "address", "Standardized_Address", "Latitude", "Longitude",
    "Google_Maps_Link", "Match_Method", "Confidence_Level",
    "Quality_Flags", "Comment",
]


# ─────────────────────────────────────────────────────────────────────────
# ZIP -> County -> State reference table
# ─────────────────────────────────────────────────────────────────────────
US_STATE_ABBR = {
    "AL", "AK", "AZ", "AR", "CA", "CO", "CT", "DE", "FL", "GA", "HI", "ID", "IL", "IN", "IA",
    "KS", "KY", "LA", "ME", "MD", "MA", "MI", "MN", "MS", "MO", "MT", "NE", "NV", "NH", "NJ",
    "NM", "NY", "NC", "ND", "OH", "OK", "OR", "PA", "RI", "SC", "SD", "TN", "TX", "UT", "VT",
    "VA", "WA", "WV", "WI", "WY", "DC", "PR",
}

ZIP_TO_REGION: dict = {}
STATE_NAME_TO_CODE: dict = {}


def load_zip_reference(path_or_buffer) -> dict:
    """Loads the ZIP -> state/county reference table from USZIP.xlsx (or an
    uploaded file-like object with the same three columns: Zipcosde,
    State Code, State Name, County). Populates the module-level
    ZIP_TO_REGION / STATE_NAME_TO_CODE dicts and returns a small summary
    dict for display in the app."""
    global ZIP_TO_REGION, STATE_NAME_TO_CODE
    zipref = pd.read_excel(path_or_buffer)
    zipref["zip5"] = zipref["Zipcosde"].astype(str).str.zfill(5)
    ZIP_TO_REGION = zipref.set_index("zip5")[["State Code", "County"]].to_dict("index")
    STATE_NAME_TO_CODE = dict(zip(zipref["State Name"].str.upper(), zipref["State Code"]))
    return {"zip_count": len(ZIP_TO_REGION), "state_count": len(STATE_NAME_TO_CODE)}


def normalize_state_token(token: str) -> str:
    token = token.strip().upper()
    if token in US_STATE_ABBR:
        return token
    if token in STATE_NAME_TO_CODE:
        return STATE_NAME_TO_CODE[token]
    return ""


def enrich_from_zip_reference(p: dict) -> list:
    flags = []
    zip5 = p.get("zip", "")
    if zip5 and zip5 in ZIP_TO_REGION:
        ref = ZIP_TO_REGION[zip5]
        ref_state, ref_county = ref["State Code"], ref["County"]
        if not p.get("state"):
            p["state"] = ref_state
            flags.append("STATE_FILLED_FROM_ZIP_REFERENCE")
        elif p["state"].upper() != ref_state.upper():
            flags.append(f"STATE_CORRECTED_FROM_ZIP_REFERENCE({p['state']}->{ref_state})")
            p["state"] = ref_state
        if not p.get("county"):
            p["county"] = ref_county
            flags.append("COUNTY_FILLED_FROM_ZIP_REFERENCE")
    return flags


# ─────────────────────────────────────────────────────────────────────────
# Character normalization (v8: full symbol stripping, - preserved)
# ─────────────────────────────────────────────────────────────────────────
# Deleted outright, no space left behind -- reconstructs the original word
# correctly for the most common corruption pattern seen in this data
# (stray symbol mid-word), and vanishes entirely for separator-like noise:
#   "Is'Land" -> "IsLand" (matches "Island" case-insensitively)
#   "123 Main St #5" -> "123 Main St 5"
# NOTE: - stays OUT of this list -- needed intact for zip+4 (12345-6789),
# hyphenated house numbers (J-132), and compound place names (Winston-Salem).
DELETE_CHARS = set("'\"`~^*$%!?©®™…&/+#")

# Replaced with a space -- these normally separate distinct words; deleting
# them outright would merge unrelated text together
SPACE_CHARS = set("()[]{}<>:;\\|@_=")

# Deliberately NOT touched: - (zip+4, hyphenated numbers, compound place
# names) and , (the CSV field separator)


def normalize_special_characters(raw: str):
    """Runs on the whole raw address string, before comma-splitting."""
    flags = []
    text = raw

    new = re.sub(r"\(\s*[^)]*\)", "", text)
    if new != text:
        flags.append("PARENTHETICAL_NOTE_REMOVED")
        text = new

    if any(c in text for c in DELETE_CHARS):
        text = "".join(c for c in text if c not in DELETE_CHARS)
        flags.append("STRAY_SYMBOL_DELETED")

    if any(c in text for c in SPACE_CHARS):
        text = "".join(" " if c in SPACE_CHARS else c for c in text)
        flags.append("SEPARATOR_SYMBOL_REPLACED")

    return re.sub(r"\s+", " ", text).strip(), flags


COMPOUND_SPLIT = re.compile(r"\s+(?:and|&|\+|/)\s+", re.IGNORECASE)

STREET_TYPES = (
    r"Rd|St|Ave|Dr|Ln|Ct|Blvd|Pl|Way|Cir|Ter|Pkwy|Hwy|Trl|Sq|"
    r"Road|Street|Avenue|Drive|Lane|Court|Boulevard|Place|Circle|"
    r"Terrace|Parkway|Highway|Trail|Square"
)

EXPLICIT_UNIT_PATTERN = re.compile(
    r"^(?P<base>.+?)\s+(?:APT|UNIT|STE|SUITE|FL|FLOOR|BLDG|BUILDING|RM|ROOM|#)\.?\s*([A-Za-z0-9\-]+)$",
    re.IGNORECASE,
)
BARE_UNIT_ALNUM_PATTERN = re.compile(
    rf"^(?P<base>.+\b(?:{STREET_TYPES}))\s+([A-Za-z]{{0,2}}\d+[A-Za-z0-9]{{0,3}})\.?$",
    re.IGNORECASE,
)
BARE_UNIT_LETTER_PATTERN = re.compile(
    rf"^(?P<base>.+\b(?:{STREET_TYPES}))\s+([A-Za-z])\.?$", re.IGNORECASE
)

ZIP_PATTERN = re.compile(r"^\d{5}(-\d{4})?$")
BARE_ZIP_PATTERN = re.compile(r"^\d{3,5}(-\d{4})?$")


def normalize_zip_field(field: str):
    field = field.strip()
    if ZIP_PATTERN.match(field):
        return field, True, []
    if re.match(r"^\d{1,4}$", field):
        return field.zfill(5), True, ["ZIP_LEADING_ZERO_RESTORED"]
    m = re.match(r"^(\d{1,5})-(\d{1,5})$", field)
    if m:
        first = m.group(1)
        return first.zfill(5), True, ["ZIP_LEADING_ZERO_RESTORED", "ZIP_RANGE_AMBIGUOUS_FIRST_SEGMENT_USED"]
    return "", False, []


def clean_text(raw: str):
    text, char_flags = normalize_special_characters(raw)
    text = re.sub(r",(?!\s)", ", ", text)
    text = re.sub(r"\s+", " ", text).strip()
    text = text.strip(" ,;")
    return text, char_flags


def strip_unit(street: str):
    m = EXPLICIT_UNIT_PATTERN.match(street)
    if m:
        return m.group("base").strip(), m.group(2)
    m = BARE_UNIT_ALNUM_PATTERN.match(street)
    if m:
        return m.group("base").strip(), m.group(2)
    m = BARE_UNIT_LETTER_PATTERN.match(street)
    if m:
        return m.group("base").strip(), m.group(2)
    return street, None


def get_street_candidates(street: str):
    segments = [s.strip() for s in COMPOUND_SPLIT.split(street) if s.strip()]
    candidates = []
    for seg in segments:
        candidates.append(seg)
        base, unit = strip_unit(seg)
        if unit:
            candidates.append(base)
    return candidates, (len(segments) > 1)


def looks_structured(parts: list) -> bool:
    if len(parts) < 3:
        return False
    tail = parts[-1].strip()
    _, is_zip_like, _ = normalize_zip_field(tail)
    if is_zip_like:
        return True
    return tail.upper() in US_STATE_ABBR


def detect_partial_shape(cleaned: str):
    parts = [p.strip() for p in cleaned.split(",") if p.strip()]

    if len(parts) == 1 and BARE_ZIP_PATTERN.match(parts[0]):
        zip5, _, zip_flags = normalize_zip_field(parts[0])
        return {"street": "", "city": "", "state": "", "county": "", "zip": zip5,
                "candidates": [], "is_intersection": False, "is_po_box": False,
                "parse_method": "bare_zip", "quality_flags": ["INPUT_IS_ZIP_ONLY"] + zip_flags}

    if len(parts) == 1:
        state_code = normalize_state_token(parts[0])
        if state_code:
            return {"street": "", "city": "", "state": state_code, "county": "", "zip": "",
                    "candidates": [], "is_intersection": False, "is_po_box": False,
                    "parse_method": "bare_state", "quality_flags": ["INPUT_IS_STATE_ONLY"]}

    if len(parts) == 2:
        state_code = normalize_state_token(parts[1])
        if state_code:
            return {"street": "", "city": parts[0], "state": state_code, "county": "", "zip": "",
                    "candidates": [], "is_intersection": False, "is_po_box": False,
                    "parse_method": "city_state_only", "quality_flags": ["INPUT_HAS_NO_STREET"]}

    return None


def parse_via_csv_structure(raw: str) -> dict:
    parts = [p.strip() for p in raw.split(",")]
    flags = []
    idx = len(parts) - 1

    zip_code = ""
    if idx >= 0:
        repaired, is_zip_like, zip_flags = normalize_zip_field(parts[idx])
        if is_zip_like:
            zip_code = repaired
            flags.extend(zip_flags)
            idx -= 1
        else:
            flags.append("MISSING_ZIP")

    county = ""
    state = ""
    if idx >= 0 and parts[idx].upper() in US_STATE_ABBR:
        state = parts[idx]
        idx -= 1
    elif idx >= 1 and parts[idx - 1].upper() in US_STATE_ABBR:
        county = parts[idx]
        state = parts[idx - 1]
        idx -= 2
    elif idx >= 0:
        state = parts[idx]
        flags.append("STATE_FORMAT_UNUSUAL")
        idx -= 1

    city = parts[idx] if idx >= 0 else ""
    idx -= 1
    street = ",".join(parts[:idx + 1]).strip() if idx >= 0 else ""

    candidates, is_intersection = get_street_candidates(street)
    if is_intersection:
        flags.append("INTERSECTION_ADDRESS")
    elif strip_unit(street)[1]:
        flags.append("HAS_UNIT_SUITE")
    if not city or not state:
        flags.append("MISSING_CITY_OR_STATE")

    return {
        "street": street, "city": city, "state": state, "zip": zip_code,
        "county": county, "candidates": candidates,
        "is_intersection": is_intersection, "is_po_box": False,
        "parse_method": "csv_structured", "quality_flags": flags,
    }


def parse_via_usaddress(raw: str) -> dict:
    flags = []
    try:
        tagged, addr_type = usaddress.tag(raw)
    except usaddress.RepeatedLabelError:
        flags.append("PARSE_AMBIGUOUS_RAW_FALLBACK")
        zip_m = re.search(r"\b\d{5}(-\d{4})?\b", raw)
        state_m = re.search(r"\b([A-Z]{2})\b", raw)
        return {
            "street": raw, "city": "", "state": state_m.group(1) if state_m else "",
            "zip": zip_m.group(0) if zip_m else "", "county": "",
            "candidates": [raw], "is_intersection": False, "is_po_box": False,
            "parse_method": "usaddress_failed", "quality_flags": flags,
        }

    city = tagged.get("PlaceName", "")
    state = tagged.get("StateName", "")
    zip_c = tagged.get("ZipCode", "")

    if addr_type == "Intersection":
        s1 = " ".join(filter(None, [tagged.get("StreetName"), tagged.get("StreetNamePostType")]))
        s2 = " ".join(filter(None, [tagged.get("SecondStreetName"), tagged.get("SecondStreetNamePostType")]))
        candidates = [c for c in [s1, s2] if c]
        flags.append("INTERSECTION_ADDRESS")
        return {
            "street": f"{s1} & {s2}", "city": city, "state": state, "zip": zip_c,
            "county": "", "candidates": candidates, "is_intersection": True,
            "is_po_box": False, "parse_method": "usaddress", "quality_flags": flags,
        }

    if addr_type == "PO Box":
        flags.append("PO_BOX")
        return {
            "street": "", "city": city, "state": state, "zip": zip_c, "county": "",
            "candidates": [], "is_intersection": False, "is_po_box": True,
            "parse_method": "usaddress", "quality_flags": flags,
        }

    number = tagged.get("AddressNumber", "")
    predir = tagged.get("StreetNamePreDirectional", "")
    name = tagged.get("StreetName", "")
    ptype = tagged.get("StreetNamePostType", "")
    postdir = tagged.get("StreetNamePostDirectional", "")
    unit_id = tagged.get("OccupancyIdentifier", "")

    clean_street = " ".join(filter(None, [number, predir, name, ptype, postdir])).strip()

    if not city or not state:
        flags.append("MISSING_CITY_OR_STATE")
    if not zip_c:
        flags.append("MISSING_ZIP")
    if unit_id:
        flags.append("HAS_UNIT_SUITE")
    if not clean_street and not city and not state:
        flags.append("UNPARSEABLE_ADDRESS")

    return {
        "street": clean_street, "city": city, "state": state, "zip": zip_c,
        "county": "", "candidates": [clean_street] if clean_street else [],
        "is_intersection": False, "is_po_box": False,
        "parse_method": "usaddress", "quality_flags": flags,
    }


def parse_address(raw: str) -> dict:
    cleaned, char_flags = clean_text(raw)

    partial = detect_partial_shape(cleaned)
    if partial is not None:
        result = partial
    else:
        parts = [p.strip() for p in cleaned.split(",")]
        if looks_structured(parts):
            result = parse_via_csv_structure(cleaned)
        elif USADDRESS_AVAILABLE:
            result = parse_via_usaddress(cleaned)
        else:
            result = parse_via_csv_structure(cleaned)
            result["quality_flags"].append("USADDRESS_UNAVAILABLE_BEST_EFFORT")

    result["quality_flags"] = char_flags + result["quality_flags"]

    enrichment_flags = enrich_from_zip_reference(result)
    result["quality_flags"].extend(enrichment_flags)

    result["raw"] = raw
    result["cleaned_raw"] = cleaned
    result["standardized"] = ", ".join(
        filter(None, [result["street"], result["city"], result["state"], result["zip"]])
    )
    return result


# ─────────────────────────────────────────────────────────────────────────
# Nominatim geocoding primitives
# ─────────────────────────────────────────────────────────────────────────
def _nominatim_request(params: dict, user_agent: str) -> Optional[dict]:
    headers = {"User-Agent": user_agent}
    try:
        resp = requests.get(NOMINATIM_URL, params=params, headers=headers, timeout=10)
        resp.raise_for_status()
        results = resp.json()
        return results[0] if results else None
    except requests.RequestException:
        return None


def geocode_structured(street="", city="", state="", zip_code="", county="",
                        user_agent=DEFAULT_USER_AGENT) -> Optional[dict]:
    params = {"format": "json", "addressdetails": 1, "limit": 1}
    if street:
        params["street"] = street
    if city:
        params["city"] = city
    if county:
        params["county"] = county
    if state:
        params["state"] = state
    if zip_code:
        params["postalcode"] = zip_code
    if COUNTRY_FILTER:
        params["country"] = COUNTRY_FILTER
    return _nominatim_request(params, user_agent)


def geocode_freeform(full_address: str, user_agent=DEFAULT_USER_AGENT) -> Optional[dict]:
    params = {"q": full_address, "format": "json", "addressdetails": 1, "limit": 1}
    if COUNTRY_FILTER:
        params["countrycodes"] = COUNTRY_FILTER
    return _nominatim_request(params, user_agent)


def result_to_dict(result: dict, method: str, confidence: str, comment: str) -> dict:
    lat, lon = float(result["lat"]), float(result["lon"])
    return {
        "Latitude": lat, "Longitude": lon,
        "Google_Maps_Link": f"https://www.google.com/maps?q={lat},{lon}",
        "Match_Method": method, "Confidence_Level": confidence, "Comment": comment,
    }


def empty_result(comment: str) -> dict:
    return {
        "Latitude": None, "Longitude": None, "Google_Maps_Link": None,
        "Match_Method": "no_match", "Confidence_Level": "NO_MATCH", "Comment": comment,
    }


def geocode_address_multi_tier(raw_address: str, user_agent=DEFAULT_USER_AGENT,
                                sleep_secs: float = SLEEP_SECS) -> dict:
    """Full cascading geocoder: structured street match -> unit-stripped/
    intersection alt -> street w/o zip -> free-text -> city -> county ->
    state. Returns a dict with Standardized_Address, Quality_Flags, and the
    geocoding result fields (Latitude/Longitude/.../Confidence_Level)."""
    p = parse_address(raw_address)
    flags_note = f" [{', '.join(p['quality_flags'])}]" if p["quality_flags"] else ""

    if p["is_po_box"]:
        result = geocode_structured(city=p["city"], state=p["state"], zip_code=p["zip"], user_agent=user_agent)
        time.sleep(sleep_secs)
        if result:
            return {**result_to_dict(result, "po_box_city_level", "CITY_LEVEL",
                     f"PO Box address -- coordinates are the city center, not a specific building.{flags_note}"),
                    "Standardized_Address": p["standardized"], "Quality_Flags": p["quality_flags"]}

    for idx, cand in enumerate(p["candidates"]):
        result = geocode_structured(cand, p["city"], p["state"], p["zip"], user_agent=user_agent)
        time.sleep(sleep_secs)
        if result:
            comment = "Matched at exact street-address level."
            if idx > 0:
                comment = ("Matched at street level after trying an alternate form "
                           "(unit-letter stripped, or the other side of a compound/intersection address).")
            comment += flags_note
            return {**result_to_dict(result, f"structured_candidate_{idx}", "EXACT_STREET", comment),
                    "Standardized_Address": p["standardized"], "Quality_Flags": p["quality_flags"]}

    if p["candidates"] and p["zip"]:
        result = geocode_structured(p["candidates"][0], p["city"], p["state"], user_agent=user_agent)
        time.sleep(sleep_secs)
        if result:
            comment = f"Matched at street level; the ZIP code was dropped from the query to find a match.{flags_note}"
            return {**result_to_dict(result, "structured_no_zip", "STREET_LEVEL_NO_ZIP", comment),
                    "Standardized_Address": p["standardized"], "Quality_Flags": p["quality_flags"]}

    if p["candidates"] or p["cleaned_raw"]:
        result = geocode_freeform(p["cleaned_raw"], user_agent=user_agent)
        time.sleep(sleep_secs)
        if result:
            comment = ("Matched using free-text search; structured parsing did not find this. "
                       f"Verify precision manually if this address is critical.{flags_note}")
            return {**result_to_dict(result, "freeform_fallback", "FREEFORM_MATCH", comment),
                    "Standardized_Address": p["standardized"], "Quality_Flags": p["quality_flags"]}

    if p["city"] and p["state"]:
        result = geocode_structured(city=p["city"], state=p["state"], zip_code=p["zip"], user_agent=user_agent)
        time.sleep(sleep_secs)
        if result:
            comment = (f"Could not resolve the exact street. Coordinates represent the "
                       f"city center for {p['city']}, {p['state']}, not the exact address.{flags_note}")
            return {**result_to_dict(result, "city_level", "CITY_LEVEL", comment),
                    "Standardized_Address": p["standardized"], "Quality_Flags": p["quality_flags"]}

    if p["county"] and p["state"]:
        result = geocode_structured(county=p["county"], state=p["state"], user_agent=user_agent)
        time.sleep(sleep_secs)
        if result:
            comment = (f"Could not resolve city or street. Coordinates represent the "
                       f"county center for {p['county']}, {p['state']}.{flags_note}")
            return {**result_to_dict(result, "county_level", "COUNTY_LEVEL", comment),
                    "Standardized_Address": p["standardized"], "Quality_Flags": p["quality_flags"]}

    if p["state"]:
        result = geocode_structured(state=p["state"], user_agent=user_agent)
        time.sleep(sleep_secs)
        if result:
            comment = (f"Could not resolve city, county, or street. Coordinates represent "
                       f"the geographic center of {p['state']} only -- very coarse.{flags_note}")
            return {**result_to_dict(result, "state_level", "STATE_LEVEL", comment),
                    "Standardized_Address": p["standardized"], "Quality_Flags": p["quality_flags"]}

    comment = f"No coordinates could be found at any level (street, city, county, or state).{flags_note}"
    return {**empty_result(comment), "Standardized_Address": p["standardized"], "Quality_Flags": p["quality_flags"]}


def preview_row(raw: str) -> dict:
    """Used by the pre-geocoding QA preview (Button 1) -- parses only, no
    network calls."""
    p = parse_address(raw)
    return {
        "Standardized_Address": p["standardized"],
        "Parse_Method": p["parse_method"],
        "County_Detected": p.get("county", ""),
        "Candidates": " | ".join(p["candidates"]) if p["candidates"] else "(none)",
        "Quality_Flags": ", ".join(p["quality_flags"]) if p["quality_flags"] else "",
    }
