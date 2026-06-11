# Handover Filler

Fill the **Shift Handover Excel template** from the **Infrastructure Handover PDF**
without breaking the template's formatting.

## Why the AI agent failed
An LLM asked to "update the Excel" either rebuilds the file from scratch (loses
formatting) or returns it unchanged (can't mutate a binary .xlsx). Fix: split the
job — **LLM extracts data → JSON**, **this code writes the cells** into the real
template. Code preserves merged cells, styles, everything.

## Setup (once)
```
python3 -m pip install pdfplumber openpyxl
```

## Steps

**1. See the template structure** (get the exact label texts):
```
python3 inspect_template.py "Shift Handover Template.xlsx"
```
Copy the exact label strings into `FIELD_MAP` at the top of `fill_handover.py`.
For each field set `dr/dc`: value **below** label → `dr=1,dc=0`; value **right**
of label → `dr=0,dc=1`.

**2. Get the data as JSON.** Either:
- Paste `extract_prompt.txt` + the PDF into your AI agent → it returns JSON → save as `data.json`, **or**
- `python3 extract_pdf.py "Infrastructure Handover (2).pdf"` to dump the PDF and fill `data.json` by hand.

The JSON keys MUST match the `key` names in `FIELD_MAP`.

**3. Preview (dry run — saves nothing):**
```
python3 fill_handover.py --template "Shift Handover Template.xlsx" --data data.json
```
Check each line: `incidents_raised -> C7 = 12`. Wrong cell? adjust `dr/dc`.
Anchor-not-found warning? the label text doesn't match — fix it from step 1.

**4. Write the filled copy:**
```
python3 fill_handover.py --template "Shift Handover Template.xlsx" --data data.json --write
```
Output: `Shift Handover Template.filled.xlsx` — the original template is untouched.

## Example (see `examples/`)
A runnable before→after demo:
- `examples/sample_template.xlsx` — blank template (merged cells, styled headers)
- `examples/sample_data.json` — extracted data
- `examples/sample_output.xlsx` — the filled result, formatting intact

```
python3 fill_handover.py --template examples/sample_template.xlsx \
                         --data examples/sample_data.json \
                         --out examples/sample_output.xlsx --write
```

## Files
- `inspect_template.py` — dump cells + merged ranges of the template
- `extract_pdf.py` — dump PDF text + tables
- `fill_handover.py` — fill the template (label-anchored, merge-aware)
- `extract_prompt.txt` — LLM prompt that outputs strict JSON
- `data.example.json` — sample data shape
