<img width="1382" height="670" alt="tests-passed" src="https://github.com/user-attachments/assets/0fecf050-b44f-43a9-bb78-58c85f1dbf98" />
# Enterprise Data Connector — Excel + SharePoint via Microsoft Graph API

A Python pipeline for a retail analytics consultancy: connects to
transaction data in Excel (OneDrive) and regional performance data in
SharePoint lists via the Microsoft Graph API, validates and standardizes
both, merges them on a common key, and produces summary analytics.

## Step 1 — Setup and Microsoft Graph authentication

### Install

```
pip install -r requirements.txt
```

### Prerequisites (do this in the Azure Portal first)

Per the course's own walkthrough (Implementing-Graph_API.txt): register an
application in **Microsoft Entra ID** → App registrations → New
registration. Since this agent runs unattended (no signed-in user), grant it
**Application permissions** (not delegated) for `Files.ReadWrite.All` (and
`Sites.Read.All` if your tenant separates SharePoint site access), then have
an admin grant consent — it won't work without that green checkmark. Create
a client secret under Certificates & secrets and copy its **Value**
immediately; it's shown exactly once.

Copy `.env.example` to `.env` and fill in the three required values:

```
AZURE_CLIENT_ID=<Application (client) ID>
AZURE_TENANT_ID=<Directory (tenant) ID>
AZURE_CLIENT_SECRET=<the client secret Value, not the Secret ID>
```

### What was built

- **[`config.py`](config.py)** — `load_graph_config()` reads credentials
  from environment variables (never hardcoded) and validates all three
  required ones up front, raising one clear error listing exactly what's
  missing — the same pattern as `kernel_setup.py` in the earlier Semantic
  Kernel Foundation activity, so a misconfigured `.env` fails fast with a
  useful message instead of a cryptic 401 three calls later.

- **[`graph_auth.py`](graph_auth.py)** — `GraphAuthenticator`, wrapping
  MSAL's **client-credentials (app-only) flow**:
  `acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])`.
  This matches the walkthrough exactly — it specifically chose Application
  permissions "for an automated agent that runs without a user present,"
  which is the client-credentials grant, not the delegated one. The MSAL
  app instance is built once and reused (not reconstructed per call), and
  MSAL's failure mode — it returns a dict with `error`/`error_description`
  keys rather than raising — is converted into a real `GraphAuthError`
  exception, so every caller has exactly one failure path to handle.

### A real finding, tested directly: MSAL needs live network access just to construct the client

Before writing tests, I tried constructing
`msal.ConfidentialClientApplication` directly. It immediately makes an
HTTPS call to `login.microsoftonline.com` for tenant discovery — even
before any token request — and this happens regardless of the
`validate_authority` flag. Confirmed live: in this project's sandboxed test
environment (network-restricted), it fails with a `ProxyError` at
construction time, not at the token-request step.

That has a real consequence for testing: you can't test this module by
mocking `requests` underneath MSAL, because the failure happens inside
MSAL's own constructor before your code's HTTP calls even begin. The fix is
dependency injection — `GraphAuthenticator` takes an `app_factory`
parameter (defaulting to the real `msal.ConfidentialClientApplication`),
so tests can substitute a lightweight fake class instead of trying to mock
network calls three layers down.

### Tested (`test_graph_auth.py`, no real Azure tenant needed)

Run: `python test_graph_auth.py` (or `pytest -v`)

- A successful token acquisition returns the token from a fake MSAL app.
- The MSAL app instance is constructed **once** and reused across three
  calls, not rebuilt each time (which would repeat the expensive discovery
  call in real usage).
- MSAL's error-dict failure mode (bad client secret, e.g.
  `AADSTS7000215`) is correctly converted into a `GraphAuthError` carrying
  both the error code and description.
- `get_auth_headers()` produces a correctly formatted `Bearer` header.
- The default scope is Graph's `.default` application-permissions scope.

All five pass without any real Azure credentials or network access —
exactly the point of injecting `app_factory`.

## Step 2 — Data fetchers: OneDrive Excel + SharePoint lists

**[`graph_client.py`](graph_client.py)** — `GraphClient`, a thin HTTP layer
over `GraphAuthenticator`: `get()` for JSON responses, `get_content()` for
raw bytes (file downloads), and `get_all_pages()` which follows Graph's
`@odata.nextLink` until it's absent. That pagination handling matters —
without it, a fetcher reading only the first page would silently truncate
any SharePoint list or OneDrive folder past Graph's page-size limit. Every
`requests` failure mode (timeout, connection error, HTTP status, invalid
JSON body) is caught and re-raised as one `GraphRequestError`, same pattern
as Step 1's `GraphAuthError` and the earlier Research Plugin Development
activity's `DataFetchError`.

**[`onedrive_fetcher.py`](onedrive_fetcher.py)** — `OneDriveFetcher` lists
files, downloads raw content, and parses Excel files straight from bytes
(`pandas.read_excel(io.BytesIO(content))`, no temp file). It deliberately
**refuses to construct without a `user`** (UPN or object id) and always
builds `/users/{user}/drive/...` paths — never `/me/...`. That's not a
style choice: this project's auth (Step 1) is the client-credentials,
app-only flow with no signed-in user, and the course's own troubleshooting
table for this exact activity names the failure directly — *"`/me` not
supported ... Switch to `/users/`."* Baking that into the constructor means
the mistake can't happen by accident later.

**[`sharepoint_fetcher.py`](sharepoint_fetcher.py)** — `SharePointFetcher`
enumerates a site's lists, then fetches one list's items with
`$expand=fields` (Graph nests actual column data under a `fields` object;
without the expand you get bare item ids and nothing else) and returns them
as a DataFrame, one row per item.

**Tested** (`test_graph_client.py`, `test_fetchers.py` — 13 tests, no real
Azure tenant needed): every `GraphClient` failure mode against a fake
session with precise per-call control; 3-page pagination via `@odata.nextLink`
followed correctly; `get_content()` verified over an **actual local HTTP
server** (real bytes over a real socket, not just a mock) since file-content
handling was worth confirming for real, the way `DataProcessingPlugin` was
in the earlier project. `OneDriveFetcher` is confirmed to always build
`/users/...` paths and never `/me/...`; `fetch_excel_as_dataframe` is
verified end-to-end against a **real `.xlsx` file** built with `openpyxl`
and parsed back with `pandas` — not a mocked DataFrame. `SharePointFetcher`
is confirmed to unpack nested `fields` correctly and to send `$expand=fields`
on every items call.

## Step 3 — Validation, standardization, and merging

**[`data_pipeline.py`](data_pipeline.py)** follows the activity's own
sample-code sequence exactly: profile missing values, standardize dates and
numbers, verify + align merge keys, merge with explicit parameters, then
summarize. Per this activity's own Question 2 answer — *"Validate and
decide according to your analysis needs"* — nothing here silently drops or
zero-fills bad data: `standardize_dates` / `standardize_numeric` return a
`ValidationReport` alongside the cleaned data, listing exactly which rows
couldn't be parsed, so that decision stays visible and in the caller's
hands rather than buried in the pipeline.

**A real finding, not a guess**: I initially wrote `standardize_dates`
using plain `pandas.to_datetime(..., errors="coerce")`. Testing it directly
against a column mixing `"01/15/2026"` and `"2026-02-01"` — exactly the
"inconsistent date formats" problem this activity is about — showed pandas
infers **one** format from the column's first value and silently fails
every row that doesn't match it: the perfectly valid `"2026-02-01"` came
back as `NaT` too, not because it's a bad date but because pandas locked
onto the first row's `MM/DD/YYYY` shape. The fix is `format="mixed"`, which
parses each value's format independently — confirmed directly before and
after the fix, not assumed.

`align_and_merge` checks the merge keys exist on both sides (raising a
clear `ValueError` naming exactly what's missing, rather than a confusing
pandas `KeyError`), casts a mismatched key dtype to match rather than
failing the merge, and logs both the merge's timing and any unmatched-row
counts — tying into this activity's later theme (Question 4: caching /
performance monitoring) by making merge cost and data-quality gaps visible
by default, not only when someone goes looking.

**Tested** (`test_data_pipeline.py` — 8 tests): missing-value profiling;
mixed-format date standardization (including the `format="mixed"` fix,
verified against the exact failure mode it corrects); currency-formatted
number parsing (`"$1,200.50"` → `1200.50`); the combined validation report;
a merge across differently-named, differently-typed key columns
(`RegionName`/`category` dtype on one side vs `Region`/plain dtype on the
other); a clear error when a key is missing from one side entirely; inner-join
unmatched-row dropping; and the region/month/target summary aggregates.

**All 26 tests across Steps 1–3 pass together**: `python -m pytest -v`.

## Step 4 — Performance, reliability, and summary visualization

This step targets the activity's own two asks: a performance/reliability
layer around the Graph calls (caching, retry on throttling, timing/logging),
and summary analytics with a visualization — plus an end-to-end `main.py`
that ties every earlier step together into one runnable pipeline.

**[`graph_client.py`](graph_client.py)** (rewritten) adds:

- **Retry with exponential backoff on HTTP 429** (Graph throttling).
  `_request()` catches a 429 response, reads its `Retry-After` header when
  present (Graph's own preferred wait time) and otherwise backs off
  `retry_backoff_seconds * 2**attempt`, retrying up to `max_retries` times
  before finally raising `GraphRequestError`. The sleep call itself is
  injectable (`sleep_fn=time.sleep`, overridden with a fake in tests) so the
  backoff logic is verified without a test suite that actually waits
  seconds for each retry.
- **Timing and logging on every request** — each attempt's duration is
  logged, so a slow or repeatedly-throttled endpoint is visible in normal
  operation rather than only when someone profiles it by hand.
- **An in-memory response cache** — `get_cached()` / `get_all_pages_cached()`
  key on `(path, sorted(params))`, so identical calls within one process
  (e.g. the same OneDrive folder listing requested by two different
  fetchers, or a re-run within the same session) hit the cache instead of
  Graph. `clear_cache()` and `cache_size()` are exposed for callers who need
  to manage it explicitly.

**[`onedrive_fetcher.py`](onedrive_fetcher.py)** and
**[`sharepoint_fetcher.py`](sharepoint_fetcher.py)** both gained a
`use_cache: bool = False` parameter on their list-returning methods,
delegating straight to the client's cached methods — opt-in, so nothing
about Steps 1–3's tested behavior changes unless a caller asks for caching.

**[`visualize.py`](visualize.py)** (new) — `plot_monthly_sales_by_region()`,
the "sample bar graph with monthly sales totals" the activity's own demo
video shows. Built using the `dataviz` skill's procedure (form before
color): a grouped bar chart, one bar-group per month, one color per region,
using the skill's validated fixed-order categorical palette (never colors
assigned by rank), a legend whenever ≥2 regions are plotted, selective
direct value labels (only when the bar count stays legible), recessive
gridlines, and no chart-junk. The skill's interactive-layer requirements
(hover tooltips, dark-mode toggle, live filters) are explicitly out of
scope here — this is a static PNG produced by a Python script, not a web
chart, which the file's own docstring notes.

**[`data_pipeline.py`](data_pipeline.py)** gained `derive_month_key()`, which
closes a real granularity gap the earlier steps hadn't addressed: the Excel
transaction data is day-level (`"2026-01-15"`), while the SharePoint
regional-performance data is month-level (`"2026-01"`). Merging directly on
the raw date would never match anything — those are two different
granularities of the same concept, and `pandas.merge` compares values, not
concepts. `derive_month_key` collapses a standardized ISO date down to its
`YYYY-MM` prefix (and correctly leaves a missing key on rows whose date
didn't parse, so they still fail to match rather than grouping under a
fabricated month).

**[`main.py`](main.py)** (new) — the end-to-end orchestration script:
`load_sources()` tries real Graph credentials and resource ids first
(`load_graph_config()` → `GraphAuthenticator` → `GraphClient` →
`OneDriveFetcher` / `SharePointFetcher`, with `use_cache=True`), and falls
back to bundled sample data — deliberately messy in the same ways the test
suite already exercises (mixed date formats, a currency-formatted value,
one genuinely bad row) — on missing config, missing resource ids, or any
fetch-time failure. `run_pipeline()` then calls
`validate_and_standardize` → `derive_month_key` → `align_and_merge` →
`summarize` → `plot_monthly_sales_by_region`, logging timing at each stage
and overall. Run: `python main.py`.

Two bottlenecks are documented directly in `main.py`'s comments, per the
activity's explicit request ("Document at least two bottlenecks"):

1. **Graph API fetch calls are the largest fixed cost per run** — each is a
   full network round trip, and the OneDrive/SharePoint fetch sequences are
   each at least two serial calls. `get_cached()` / `get_all_pages_cached()`
   remove this cost on any *repeated* call within the same process (proven
   in `test_graph_client_step4.py`: 3 identical calls → 1 real network
   call), but the first call's network cost is unavoidable.
2. **`get_all_pages()` fetches pages serially**, one full round trip at a
   time — for a very large SharePoint list this would dominate runtime.
   Not fixed here (out of scope for this activity); a production version
   would request a larger page size via `$top` or fetch pages concurrently.

**Tested**: `python main.py` runs the full pipeline against the bundled
sample data (no real Azure tenant available in this environment) — the one
row with an unparseable date/amount is correctly flagged in the validation
report and excluded from the merge rather than silently corrupting a total;
the rendered chart was visually inspected (not just assumed correct from
the code) and shows the expected two-region, three-month grouped bars with
March/East correctly absent, since that's the row with the bad data.
`test_graph_client_step4.py` (7 new tests) covers the retry/backoff/cache
logic directly: 429 → retry → success, `Retry-After` header honored,
giving up after `max_retries`, and cache hit/miss behavior including
per-parameter cache keys. One test each was added to `test_fetchers.py` and
`test_data_pipeline.py` for the new cache pass-through and
`derive_month_key`, respectively.

**All 35 tests across Steps 1–4 pass together**: `python -m pytest -v`.

## Project layout

```
requirements.txt
requirements-dev.txt
pytest.ini
.env.example
.gitignore
config.py                Step 1 — env-var credential loading, validated up front
graph_auth.py             Step 1 — MSAL client-credentials auth, testable via app_factory
test_graph_auth.py        Step 1 — auth tests against fake MSAL apps
graph_client.py           Step 2/4 — Graph HTTP client: get / get_content / get_all_pages,
                          plus Step 4's retry-on-429, timing/logging, and in-memory cache
onedrive_fetcher.py       Step 2/4 — OneDrive file listing + Excel-from-bytes (app-only, /users/ not /me/), optional caching
sharepoint_fetcher.py     Step 2/4 — SharePoint list enumeration + item retrieval, optional caching
test_graph_client.py      Step 2 — GraphClient tests (fake session + one real local HTTP server)
test_graph_client_step4.py Step 4 — retry/backoff and cache tests
test_fetchers.py          Step 2/4 — fetcher tests (real .xlsx round trip + fake SharePoint items + cache reuse)
data_pipeline.py          Step 3/4 — validate/standardize/merge/summarize, plus derive_month_key
test_data_pipeline.py     Step 3/4 — pipeline tests against deliberately messy sample data
visualize.py              Step 4 — monthly-sales-by-region bar chart (dataviz skill)
main.py                   Step 4 — end-to-end pipeline: fetch (or sample fallback) -> validate -> merge -> summarize -> chart
```
<img width="1382" height="670" alt="tests-passed" src="https://github.com/user-attachments/assets/c79bdb89-4834-4301-9c35-fc3c1fa7bd4d" />
