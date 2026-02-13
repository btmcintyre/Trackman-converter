<!-- Copilot / AI agent instructions for TrackmanConverter -->
# Quick Orientation

TrackmanConverter is a small Windows-focused GUI tool that:
- Finds a user's TrackMan report from Chrome history/cookies.
- Downloads the report JSON via TrackMan API.
- Converts the JSON into a formatted Excel workbook (OpenPyXL).

Key entrypoints and files:
- `trackman_gui_app.py` — the Tk/CustomTkinter GUI and conversion orchestration.
- `trackman_api.py` — Chrome-history parsing, report lookup, and download (`download_report`, `get_all_report_ids_from_chrome`, `fetch_report_metadata`).
- `trackman_auth.py` — cookie/token extraction and token persistence (`get_saved_token`, `save_token`, `extract_token_from_chrome`, `login_via_browser`).
- `trackman_gui_app.spec` — PyInstaller spec used to build the Windows executable.
- `trackman_full_report.json` — sample/last-downloaded JSON saved by `download_report`.
- `trackman_token.txt` — token file location (created either next to the frozen exe or in repo during development).

Design & data flow (what to know first)
- The GUI calls `trackman_auth` to obtain a bearer token (tries Chrome cookies, falls back to manual paste). If frozen with PyInstaller, `TOKEN_FILE` is written next to the executable; in dev mode it's in the repo base.
- With the token, `trackman_api` scans a local Chrome History DB copy to find recent TrackMan report URLs, then calls the TrackMan report API to download the JSON to `trackman_full_report.json`.
- `trackman_gui_app.convert_json_to_excel()` reads that JSON, builds per-club sheets and an "All Data" sheet, then prompts the user to save an `.xlsx` via a standard file dialog.

Platform & environment notes
- Target: Windows (reads Chrome profile under `%LOCALAPPDATA%\Google\Chrome\User Data\Default`).
- The code expects Chrome's `History` and `Cookies` SQLite files; cookies extraction may fail if Chrome is running — close Chrome for reliable cookie copying.
- The GUI and build were created with PyInstaller (see `trackman_gui_app.spec` and `build/trackman_gui_app/`).

Dependencies (discoverable from imports)
- UI: `customtkinter` (plus built-in `tkinter`)
- Data: `pandas`, `openpyxl`
- HTTP/IO: `requests`
- Stdlib: `sqlite3`, `shutil`, `tempfile`, `pathlib`, `json`

Common developer workflows
- Run the GUI locally (see prints for progress):
```
python trackman_gui_app.py
```
- Build a frozen Windows executable (requires PyInstaller installed):
```
pyinstaller trackman_gui_app.spec
```
- If Chrome cookie extraction fails, run the app and paste the bearer token when prompted; the token will be saved to `trackman_token.txt` (path depends on frozen vs dev mode).

Project-specific conventions & patterns
- Token persistence: `trackman_auth.TOKEN_FILE` resolves to the running exe's parent directory when frozen; otherwise it uses the repository base (`BASE_DIR`). Tests or automation should account for this path variation.
- Chrome access: code always copies Chrome DBs to a temporary file before reading. Tests should mock `shutil.copyfile` / `sqlite3.connect` or provide a sample DB.
- Output: converted Excel files are saved via the user's file dialog; a `trackman_full_report.json` file is kept at repo root after downloads — useful for debugging conversions.
- Logging: the code uses `print()` for quick feedback; run from a console to see these runtime messages for debugging.

Quick code examples to reference
- Token read/save: see `trackman_auth.get_saved_token()` and `trackman_auth.save_token(token)`.
- Download flow: `trackman_gui_app.handle_cloud()` -> `trackman_api.get_all_report_ids_from_chrome()` -> `trackman_api.fetch_report_metadata()` -> `trackman_api.download_report()` -> `convert_json_to_excel()`.

Safety and permissions
- Cookie extraction reads Chrome profile files — do not commit actual cookie/token files. `trackman_token.txt` may contain bearer tokens; treat as secrets if present.

- If you need more
- If something in these instructions is unclear or you want an expanded section (build CI, a `requirements.txt`, or a sample Chrome DB for tests), tell me which area to expand and I'll update this file.

**Testing & CI (practical tips)**

- Local smoke tests (fast, non-GUI): exercise the conversion logic without showing dialogs by importing the conversion function and using the sample JSON:

```
python -c "import json; from trackman_gui_app import build_workbook_per_club; d=json.load(open('trackman_full_report.json')); wb=build_workbook_per_club(d); wb.save('out.xlsx')"
```

- Unit tests: focus on `trackman_api` and `trackman_auth` by mocking file access and `sqlite3` results. Example targets:
	- `get_all_report_ids_from_chrome()` — provide a small copied Chrome `History` SQLite (or mock `sqlite3.connect`) and assert returned IDs.
	- `extract_token_from_chrome()` — test with a local temporary cookie DB or mock the file read/SQLite cursor.
	- `build_workbook_per_club()` — load `trackman_full_report.json`, verify sheets exist and expected column headers.

- Headless CI notes: avoid GUI calls (file dialogs, Tk mainloop) in CI. Import helpers directly (`build_workbook_per_club`) rather than running `trackman_gui_app`'s `mainloop`.

- Secrets for integration tests: if you need to call the real TrackMan API in CI, store a bearer token in repository secrets (`TRACKMAN_TOKEN`) and reference it in the workflow. Prefer not to do real API calls on PRs.

**Sample GitHub Actions workflow (Windows):**

Create `.github/workflows/ci.yml` with a minimal smoke-test flow. Key steps shown below:

```yaml
name: CI
on: [push, pull_request]
jobs:
	smoke-test:
		runs-on: windows-latest
		steps:
			- uses: actions/checkout@v4
			- name: Setup Python
				uses: actions/setup-python@v5
				with:
					python-version: '3.11'
			- name: Install dependencies
				run: python -m pip install --upgrade pip && pip install pandas openpyxl requests
			- name: Run conversion smoke test
				run: |
					python -c "import json; from trackman_gui_app import build_workbook_per_club; d=json.load(open('trackman_full_report.json')); wb=build_workbook_per_club(d); wb.save('out_ci.xlsx')"
			- name: Upload artifact
				uses: actions/upload-artifact@v4
				with:
					name: converted-xlsx
					path: out_ci.xlsx
```

- Optional: add linting (`flake8`, `ruff`) or a PyInstaller packaging job for release builds (packaging on CI requires additional runner setup and may increase runtime).

**Testing patterns & mocks (project-specific)**

- Chrome DB access is always done by copying the DB to a temp file then opening it. Tests should mock `shutil.copyfile` or provide a small sample DB in `tests/fixtures/` and point the functions to it by monkeypatching `Path.home()` or overriding the path via a helper.
- Cookie extraction decodes cookie values and reads the `cookies` table. Provide a minimal SQLite fixture with `name='appsession'` row to simulate a saved token.

**CI security & secrets**

- If CI needs to call `download_report`, provide `TRACKMAN_TOKEN` as a repository secret and use it in the workflow step as an env var. Do not commit `trackman_token.txt` or any real tokens.

-- End of file
-- End of file
