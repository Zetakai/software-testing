# Software Testing — Studi Kasus SauceDemo

Manual software testing case study for the course **Pengujian Perangkat Lunak**
(Topic: *Test Execution & Test Reporting*). Target under test:
[SauceDemo / Swag Labs](https://www.saucedemo.com/) — a public e-commerce demo app.

**Tim:** Zaki & Galih

---

## Scope

Black-box manual testing across 5 modules — Login, Home Page, Cart, Checkout,
Footer — covering positive, negative, edge-case, and responsiveness scenarios.

| Metric | Value |
|--------|-------|
| Total test cases | **99** |
| Passed | **82** |
| Failed | **17** |
| Pass rate | **82.8%** |
| Bugs reported | **17** (0 critical / 0 high severity, 4 medium, 13 low) |
| Release decision | **Conditional GO** |

The 3 highest-priority bugs concern the core transaction flow: cannot add an item
more than once, missing delivery-address field, and missing payment-method
selection in checkout.

---

## Repository contents

| File | Description |
|------|-------------|
| `STUDI KASUS PENGUJIAN PERANGKAT LUNAK.pdf` | Assignment brief (required format & rubric) |
| `SauceDemo (Test Cases).xlsx` | Test Case Document + Test Execution Result (per PDF §C.1–C.2) |
| `SauceDemo (Bug Report).xlsx` | Defect Report (per PDF §C.3) |
| `SauceDemo_Test_Report.pptx` | Presentation deck (13 slides) |
| `SauceDemo_Test_Report.pdf` | PDF render of the deck |
| `gen_ppt.py` | Generator for the presentation deck (python-pptx) |
| `reformat_xlsx.py` | Reformats the workbooks to the PDF field specs (openpyxl) |
| `run_tests.py` | Selenium harness — live verification of key scenarios |
| `live_results.json` | Output of the latest live test run |

### Document formats (per the studi-kasus PDF)

**Test Cases:** `ID Test | Nama Skenario | Langkah Pengujian | Expected Result | Hasil Aktual | Status | Catatan Pengujian`
Status values: `PASS · FAIL · BLOCKED · NOT EXECUTED · RETEST`.

**Bug Report:** `Bug ID | Severity | Priority | Deskripsi Defect | Langkah Reproduksi | Expected Result | Actual Result`.

---

## Reproduce

```bash
# Regenerate the deck (PPTX)
python gen_ppt.py

# Reformat the workbooks to the PDF spec
python reformat_xlsx.py

# Run the live Selenium verification (needs Chrome + selenium)
python run_tests.py
```

Dependencies: `python-pptx`, `openpyxl`, `selenium` (Chrome required for `run_tests.py`).

---

## Notes

- The live Selenium run (`run_tests.py`) independently verified core flows on
  `saucedemo.com`; results are in `live_results.json`.
- SauceDemo intentionally ships buggy accounts (`problem_user`, `error_user`,
  etc.) as testing practice material.
