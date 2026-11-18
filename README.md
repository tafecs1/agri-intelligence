# Agri Intelligence India — GitHub runner

`agri_colab_auto.py` is the **same notebook code** as Colab (CELL 1 → CELL 17, all logic, sources, thresholds,
scores and tabs unchanged). `run.py` executes its cells in order, exactly like Colab "Run all".
The only edits are 4 additive GitHub adapters, marked `# [GitHub]` and listed in `CHANGES_FROM_COLAB.diff`:

| Adapter | Why | Effect in Colab |
|---|---|---|
| Earth Engine key read from env `GEE_KEY` | GitHub passes secrets as environment variables | none |
| No interactive Earth Engine login when `CI` is set | no browser on GitHub | none |
| Google Sheets written by a service account into **fixed** Sheet IDs | no Colab sign-in on GitHub; fixed links for the dashboard | none (Colab upload path unchanged) |
| Optional `AGRI_STATES` test filter | quick test runs | none (empty = all India) |

## One-time setup (≈15 min)
1. **Create the repo** (public is fine: keys live in Secrets; data files are not attached to runs), upload these files (never upload the JSON key — `.gitignore` blocks `*.json`).
2. **Create 2 empty Google Sheets in your Drive:** "Agri Intelligence India" (all 24 main + validation tabs) and
   "Agri Trend Module" (tabs A-G). Share both with the service account e-mail (`client_email` in your key) as **Editor**.
   Copy each Sheet ID (the part between `/d/` and `/edit` in the link). Output = these 2 Sheets, overwritten every Monday.
   (Two Sheets because everything together is ~12 million cells and one Google Sheet holds max 10 million.)
3. **Repo → Settings → Secrets and variables → Actions → New repository secret:**

| Secret | Value | Needed |
|---|---|---|
| `GEE_KEY` | whole service-account JSON text | yes |
| `SHEET_ID_MAIN`, `SHEET_ID_TREND` | Sheet IDs | yes |
| `SHEET_ID_VALIDATION`, `SHEET_ID_TREND_SOWING` | safety net only: used if one workbook alone ever passes 9.5M cells (today: main ~7.2M, trend ~5.0M) | optional |
| `SHEETS_KEY` | another service-account JSON for Drive (else `GEE_KEY` is used) | optional |
| `DATA_GOV_IN_KEY`, `GEMINI_API_KEY`, `OPENAI_API_KEY`, `ANTHROPIC_API_KEY` | same as Colab | optional |

4. In Google Cloud for the key's project, enable **Google Drive API** (Earth Engine API is already on).

## Running
- **Test first:** Actions → *Agri Intelligence* → *Run workflow* → states = `TAMIL NADU` (≈ 20–40 min).
- **Full run:** leave inputs empty, or wait for the weekly schedule: **every Monday 02:00 IST, whole code (CELL 1-17)**.
- **Main workbook only:** cells = `1-14,17`. **Trend only needs the main cells first**, so run all.
- Links appear in the run's **Summary** page; the run log is in **Artifacts** (14 days). Data files stay in the work cache, not attached (public repo).
- If a run hits the 6-hour limit, re-run it the same day: downloads and Earth Engine checkpoints are cached and it resumes.

## Local run (any PC)
```
pip install -r requirements.txt
set GEE_KEY=<json text>   (Windows)   |   export GEE_KEY='<json text>'   (Mac/Linux)
python run.py
```
