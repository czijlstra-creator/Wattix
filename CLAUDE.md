# Wattix Automation — Project Brief

## What this project is
A Streamlit app (`app.py`) with two tabs:
1. **Excel Generator** — takes Wattix CSV/XLSX files and produces a formatted Bluepeak Excel report
2. **PPT Generator** — takes that Excel output and fills a PowerPoint template with scenario data

The app is at: `Documents\GitHub\Wattix\app.py`
Source/reference scripts are in: `Attachments\Wattix automation\`

---

## Current scope — PPT Generator only

**Only work on the PPT Generator tab unless explicitly asked otherwise.**

The Excel Generator is working. Do not touch it.

### PPT output rules (non-negotiable)
- Output = **scenario tech slides only** — one per scenario, no more
- No observation slides
- No cover pages
- No appendix or closing slides
- The `keep_slides` set must only contain `tech_slide` from each active pair

### Active bug (as of 2026-04-15)
The generated PPT still shows "can't read" in PowerPoint desktop.
- The file structure is verified identical to the original working script (`wattix_to_pptx.py`)
- Both produce: 1 slide (slide6.xml), 35 orphaned Content_Types entries, same presentation.xml
- Root cause not yet confirmed — next step is ruling out browser download corruption vs file bug

---

## Key technical facts

- `save_to_bytes()` is a pure zip write — no XML manipulation on save
- Chart external data (`c:externalData` + SharePoint OLE rels) must NOT be stripped — that was the bug in a prior session, now fixed
- `SCENARIO_SLIDE_PAIRS` maps scenario index → (tech_slide, obs_slide) filenames
- `SCENARIO_CHART_GROUPS` maps scenario index → list of 4 chart filenames
- Template has 3 scenario pairs (slides 6-11), 4 charts each (charts 3-14)
- Charts are normalised to last pair's bytes before data update so all slides share consistent formatting

---

## What NOT to do
- Do not add cover slides back to output
- Do not add observation slides to output
- Do not modify the Excel Generator unless explicitly asked
- Do not change working folders — all edits go to `mnt/Wattix/app.py`
- Do not go in circles retrying the same fix — if a fix doesn't work, diagnose before retrying
