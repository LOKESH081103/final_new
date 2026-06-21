# DHC Working Automation — Roadmap

## 1. How the current manual file actually works (reverse-engineered)

Excel keeps a cached copy of any formula that points at another file, even after
the link breaks. That cache is how this was confirmed — not guessed.

| Sheet / columns | Source | Logic |
|---|---|---|
| `Look Up` cols C:D | `To_be_Disabled.xlsb` → *Agreement Level Disble* | AGREEMENTNO → disable Status |
| `Look Up` cols F:G | `To_be_Disabled.xlsb` → *CIF Level Disable* | CIF_NO → disable Status |
| `Look Up` col A | `CIFCL_CBSL_List.xlsx` | Employee mobile numbers, used to flag receipts entered on an agent's phone |
| `Look Up` cols J:N | **Carried forward month to month** | AGREEMENTNO → CIF_NO / Zone / Sub Region / Opening Slab, ~5,200 agreements known so far |
| `DCR` tab cols 1–73 | `DCR.xlsb` Sheet1 | Pasted as-is (17,256 rows in the file you sent) |
| `DCR` tab cols 74–84 | Formulas | 11 derived columns: CIF, Unique Mob number, Mob Num vs Emp Mob Num, Mode, Status, Receipt Source, Zone, Sub Region, Slab, Ag Level cash mode, CIF LEVEL |
| `DCR` Sheet2 / `[2]Sheet2` | `DCR.xlsb` Sheet2 | 116,629-row full agreement master for Opening DPD/Slab |
| `RTGS` tab | Subset of `DCR` tab | Filtered to RTGS-mode receipts; adds Ageing/TAT/Receipt Type |
| `Receipt made summary`, `RTGS Summary`, `Cash Mode Validat Summary`, `Delay in RCPTING Summary` | Pure aggregation | COUNTIFS/SUMIFS/pivot off DCR & RTGS — fully mechanical |
| `RCPT CXN` | Manual log | Free-text remarks column — a judgment call, not a formula |

**Mapping tables copied from the workbook's `Sheet3` tab** (small, static, reused
in `etl.py` as Python dicts): TAT bucket boundaries, Receipt Type normalization,
Mode of Payment normalization, Status code expansion (B/C/D/NA/X), Receipt
Source normalization.

## 2. Open questions to confirm with your mam before going live

1. **CIF mapping for new agreements.** The `Look Up` master only "knows" ~5,200
   agreements. Any agreement that hasn't appeared in a DCR before has no CIF
   number, no Zone, no Sub Region in any of the 4 files — that information
   isn't in this data at all. The prototype carries forward whatever the
   previous month's workbook already knew and clearly flags what's still
   missing, but someone needs to confirm **where she currently gets that
   mapping from** (a separate LMS export? a manual lookup tool?) so it can be
   automated too, or accepted as a small manual top-up step each month.
2. **RTGS tab scope.** In the file you sent, the RTGS tab only has 100 rows
   covering June 1–3, while DCR covers the whole month (17,256 rows, 1,879 of
   them RTGS-mode). Is RTGS meant to be the *full* month's RTGS receipts
   (what the prototype currently does), or a smaller working list she
   prunes once a receipt clears?
3. **RCPT CXN remarks.** This sheet is a judgment log — the app can
   pre-populate every cancelled receipt automatically, but the `Remarks`
   column is left blank by design for her to fill in.
4. **Receipt made summary exact layout.** The original has two side-by-side
   pivot blocks (Updated/Pending vs Cancelled, with slightly different
   sub-columns). The prototype produces one combined Status × Mode × sub-status
   table with the same numbers — worth a side-by-side check against a real
   day's output before this replaces the manual version.

## 3. Phased rollout plan

**Phase 0 — Validate the logic (1–2 days)**
Run the attached app on 2–3 past months' files, diff its output against what
mam already produced by hand for those months. Resolve the open questions
above. This is the highest-leverage step — everything downstream depends on
trusting the numbers.

**Phase 1 — Harden the ETL (close gaps found in Phase 0)**
- Wire in the real source of the CIF/Zone/Sub Region mapping if one exists.
- Match the exact pivot layout of `Receipt made summary` if the simplified
  version isn't acceptable.
- Add data-quality checks: duplicate receipt numbers, missing AGREEMENTNO,
  date parsing failures — currently silent `NaN`s, should become visible
  warnings in the UI.

**Phase 2 — Productionize the Streamlit app**
- Add a "compare to last month" sanity-check screen (row counts, total
  AMOUNTPAID) before allowing download — catches a wrong file upload early.
- Add an audit log (who ran it, when, with which files) if more than one
  person will use this.
- Style the output workbook to match the original's column widths/colors if
  it needs to look identical for downstream consumers.

**Phase 3 — Deploy**
- Easiest: run `streamlit run app.py` on a shared machine / internal server,
  or Streamlit Community Cloud if the data isn't sensitive enough to need to
  stay on-prem (it likely is, given customer CIF/mobile data — prefer an
  internal server).
- mam (or whoever runs it) opens a browser tab, uploads the 4 files, clicks
  Process, downloads the result. No Excel formulas to maintain by hand again.

**Phase 4 — Retire the manual workbook**
Once Phase 0–3 numbers match for a few consecutive months, switch over fully
and keep the old `DHC_Working.xlsx` only as the monthly "previous working
file" input for the CIF carry-forward.

## 4. What's in this delivery

- `etl.py` — the transformation pipeline, **tested end-to-end against your
  actual 4 files** (not synthetic data). Every loader and builder function
  ran successfully on the real DCR/disable-list/employee/prior-workbook data.
- `app.py` — the Streamlit UI: 4 upload slots, a Process button, a preview
  tab per output sheet, and a download button. Boots clean with no errors.
- `requirements.txt` — `pip install -r requirements.txt`, then
  `streamlit run app.py`.

## 5. Known simplifications in this first pass

- Output sheet values are written as plain numbers, not live Excel formulas —
  intentional. The whole point of automating this is that pandas computes
  the joins/aggregations instead of VLOOKUP; baking formulas back in would
  just reintroduce the fragility (the workbook already has at least one
  broken/orphaned formula from a copy-paste mistake, found during this
  analysis — col N of `Look Up` references columns that don't exist on that
  sheet).
- `Cash Mode Validat Summary` and `RCPT CXN` column labels are a best-effort
  match to the original headers — confirm exact wording with mam if this
  feeds into a report template elsewhere.
