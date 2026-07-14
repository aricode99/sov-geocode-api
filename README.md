# SOV Geocode API -- direct Excel integration (no add-in)

A thin FastAPI wrapper around `geocoder_core.py` (unchanged from the
Streamlit app), deployed straight from this repo, callable from plain
Excel formulas or Power Query -- no client install, no add-in, no
PyInstaller.

## ⚠️ Read this first -- Nominatim blocks cloud/datacenter IPs

While building this, a direct test from a cloud sandbox against
`nominatim.openstreetmap.org` returned:

```
HTTP 403: Access denied. See https://operations.osmfoundation.org/policies/nominatim/
```

This is Nominatim's public server actively blocking a cloud-hosted IP
range, **not** a bug in this code -- confirmed by testing the exact same
request directly against Nominatim's API, outside our own code entirely.

This matters because **Render, Fly.io, and Azure App Service are also
cloud IPs**. Whether your specific host's outbound IP is on Nominatim's
block list is something you can only confirm by actually deploying and
testing -- it isn't predictable from here. Before relying on this in
production:

1. Deploy to your chosen host and run one real `/geocode` call against a
   known address. If you get real coordinates back, you're clear.
2. If you get `NO_MATCH` on addresses that obviously should resolve (the
   same silent failure mode this sandbox hit -- our code treats a 403
   the same as "no results found," so it won't look like an error),
   that's the IP-blocking issue, not a code or data problem.
3. **Mitigations if your host is blocked:**
   - Self-host Nominatim (Docker image available) on infrastructure with
     an IP not subject to the public policy -- more infrastructure to
     run, but removes the dependency on the shared public instance
     entirely.
   - Switch to a commercial geocoder with a proper API key and cloud-
     friendly terms (Mapbox, Google, HERE, etc.) -- means changing
     `geocoder_core.py`'s Nominatim calls, a real but contained change.
   - Fall back to the **local Windows service** architecture from the
     earlier design doc (Section A1/A2) -- calls go out from each user's
     own machine/corporate network IP, which is far less likely to be
     blocked than a datacenter range, at the cost of requiring a client
     install again.

This repo's tests (`tests/test_api.py`) deliberately mock the geocoder
call rather than hitting live Nominatim, specifically so CI doesn't fail
due to this IP-blocking issue rather than an actual code problem.

## Setup

### 1. Deploy to Render

1. Push this repo to GitHub.
2. On [render.com](https://render.com), **New -> Web Service**, connect
   this repo. Render auto-detects `render.yaml`.
3. In the service's **Environment** tab, set:
   - `GEOCODE_API_KEY` -- generate any long random string (this gates who
     can call your API).
   - `NOMINATIM_CONTACT_EMAIL` -- required by Nominatim's usage policy.
4. Once deployed, note your service URL (e.g.
   `https://sov-geocode-api.onrender.com`) and test:
   ```
   curl "https://sov-geocode-api.onrender.com/geocode?address=1282+Pickett+Street,+Charleston,+SC,+29412&key=YOUR_KEY"
   ```

### 2. Wire up GitHub Actions auto-deploy

1. In Render's dashboard, under the service's **Settings**, find
   **Deploy Hook** and copy the URL.
2. In your GitHub repo, **Settings -> Secrets and variables -> Actions**,
   add:
   - `RENDER_DEPLOY_HOOK_URL` -- the URL from step 1.
   - `API_HEALTH_URL` -- `https://sov-geocode-api.onrender.com/health`.
3. Push to `main` -- `.github/workflows/deploy.yml` runs the test suite
   first, and only triggers the Render deploy hook if tests pass.

### 3. Free-tier caveats worth knowing

- Render's free plan **spins down after inactivity** -- the first
  request after idle can take 30-50 seconds while it wakes up. Excel
  formulas relying on `WEBSERVICE()` may show `#VALUE!` if that exceeds
  Excel's internal timeout; retrying the formula (F9) after the service
  is awake resolves it.
- Render's free plan has an **ephemeral filesystem** -- the SQLite cache
  (`geocode_cache.sqlite3`) resets on every redeploy and on every
  spin-down/spin-up cycle. This only costs you re-geocoding (extra
  Nominatim calls), not correctness -- but if you want a durable cache
  across restarts, that needs Render's paid persistent disk add-on or an
  external store (e.g. a small hosted Postgres/Redis).

## Calling it from Excel

### Option 1 -- `LAMBDA()` custom functions (Excel 365 only)

Excel has no native JSON parser in plain formulas. The commonly-cited
`FILTERXML()` trick breaks on this API's actual response shape, because
fields like `standardized` and `comment` contain commas and colons
(e.g. `"1282 Pickett Street, Charleston, SC, 29412"`) -- naive
JSON-to-XML conversion by splitting on commas mangles those values.

Instead, the formulas below extract each field by locating its JSON key
directly and reading up to the next unescaped delimiter -- validated
against this API's real response format, including the `null`
lat/lon case for unmatched addresses:

**Set these up once via Formulas -> Name Manager -> New (or Name Box ->
type the name -> Ctrl+Enter with the formula bar, for a new Excel 365
workbook):**

```
Name: ApiBaseUrl
Refers to: ="https://sov-geocode-api.onrender.com"

Name: ApiKey
Refers to: ="YOUR_API_KEY_HERE"

Name: JSONSTRINGFIELD
Refers to:
=LAMBDA(json_text, field_name,
    LET(
        search_key, """" & field_name & """:""",
        start_pos, FIND(search_key, json_text) + LEN(search_key),
        end_pos, FIND("""", json_text, start_pos),
        MID(json_text, start_pos, end_pos - start_pos)
    )
)

Name: JSONNUMFIELD
Refers to:
=LAMBDA(json_text, field_name,
    LET(
        search_key, """" & field_name & """:",
        start_pos, FIND(search_key, json_text) + LEN(search_key),
        comma_pos, IFERROR(FIND(",", json_text, start_pos), LEN(json_text) + 1),
        brace_pos, IFERROR(FIND("}", json_text, start_pos), LEN(json_text) + 1),
        end_pos, MIN(comma_pos, brace_pos),
        raw_val, MID(json_text, start_pos, end_pos - start_pos),
        IF(raw_val = "null", NA(), VALUE(raw_val))
    )
)

Name: GEOCODE
Refers to:
=LAMBDA(address,
    LET(
        url, ApiBaseUrl & "/geocode?address=" & ENCODEURL(address) & "&key=" & ApiKey,
        json_text, WEBSERVICE(url),
        HSTACK(
            JSONNUMFIELD(json_text, "lat"),
            JSONNUMFIELD(json_text, "lon"),
            JSONSTRINGFIELD(json_text, "confidence"),
            JSONSTRINGFIELD(json_text, "standardized")
        )
    )
)
```

**Usage in a cell:**
```
=GEOCODE(A2)
```
Spills four values (lat, lon, confidence, standardized) into the cell
and the three cells to its right, from one formula -- same ergonomics as
the `=GEOCODE()` UDF from the add-in design, with zero installation.

**Known limitations of this approach** (be upfront with users about
these):
- If an address string itself contains a literal `"` character, the
  string-field extraction breaks (rare for real addresses, but not
  impossible with messy SOV data -- worth a sanity check upstream, same
  as the SOV Cleansing Tool already does).
- `WEBSERVICE()` has an internal ~60 second timeout and Excel does not
  pace multiple simultaneous formula recalculations -- if you paste
  `=GEOCODE()` down 200 rows and hit F9, Excel may fire many requests
  at once, which will queue up behind this API's shared rate limiter
  (~1.1s apart) and some may time out waiting their turn. **For anything
  beyond a handful of one-off lookups, use Power Query (Option 2)
  instead** -- it processes rows sequentially and doesn't have this
  pile-up risk.

### Option 2 -- Power Query (recommended for bulk/batch use)

```
let
    ApiBaseUrl = "https://sov-geocode-api.onrender.com",
    ApiKey = "YOUR_API_KEY_HERE",

    CallGeocode = (address as text) as record =>
        let
            url = ApiBaseUrl & "/geocode?address=" & Uri.EscapeDataString(address) & "&key=" & ApiKey,
            response = Json.Document(Web.Contents(url))
        in
            response,

    Source = Excel.CurrentWorkbook(){[Name="Addresses"]}[Content],
    AddedGeocode = Table.AddColumn(Source, "Geocode", each CallGeocode([Address])),
    Expanded = Table.ExpandRecordColumn(
        AddedGeocode, "Geocode",
        {"lat", "lon", "confidence", "standardized", "match_method", "comment", "cached"}
    )
in
    Expanded
```

This calls the API once per row via `Table.AddColumn`, which Power
Query evaluates sequentially by default -- naturally respecting the
API's rate limiter without the pile-up risk `WEBSERVICE()` has. Use
**Data -> Refresh All** to run/re-run the batch.

### Option 3 -- `BATCHGEOCODE` for genuinely large batches

For very large address lists, call `/batch_geocode` (comma-free,
pipe-separated to sidestep the "commas inside addresses" issue) directly
via Power Query, chunking your address list into groups of ≤100 (the
API's own limit, matching the `max_results` guardrails elsewhere in this
project):

```
let
    ApiBaseUrl = "https://sov-geocode-api.onrender.com",
    ApiKey = "YOUR_API_KEY_HERE",
    Addresses = Excel.CurrentWorkbook(){[Name="Addresses"]}[Content][Address],
    Joined = Text.Combine(Addresses, "|"),
    url = ApiBaseUrl & "/batch_geocode?addresses=" & Uri.EscapeDataString(Joined) & "&key=" & ApiKey,
    response = Json.Document(Web.Contents(url)),
    ResultsList = response[results],
    ResultsTable = Table.FromRecords(ResultsList)
in
    ResultsTable
```

## Testing this yourself

```bash
pip install -r requirements.txt -r requirements-dev.txt
pytest tests/ -v                    # mocked, no live Nominatim dependency
uvicorn main:app --reload           # run locally on http://localhost:8000
```
