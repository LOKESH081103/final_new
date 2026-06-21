# DHC Working Automation — Complete Setup & Usage Guide

## 🚀 Quick Start (5 minutes)

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Run the App
```bash
streamlit run app.py
```

Open your browser to `http://localhost:8501`

### 3. Upload & Process
- **Upload 4 files**: DCR, To be Disabled, Employee Master, Previous DHC Working
- Click **🚀 Process Files**
- Download the output Excel workbook

---

## 📊 What the App Does

**Fully automates** the manual VLOOKUP/pivot workflow your mam currently does:

1. **Loads 4 source files** (DCR receipts, compliance lists, employee mobile numbers, prior master)
2. **Builds derived columns** (CIF, Mode, Status, Receipt Source, Zone, Sub Region, Slab, disable flags)
3. **Generates 5 professional pivot-style summaries** with merged headers, subtotals, & formatting
4. **Outputs a clean Excel workbook** ready to share or use downstream

### Processing Time
- **Full pipeline**: ~40–50 seconds (loading + transforms + Excel write)
- **No manual formulas** — all logic is deterministic Python, not fragile VLOOKUP

---

## 📋 Input Files

### 1. **DCR.xlsb** (Daily Collection Report)
- **Sheet1**: Receipt-level raw data (17,000–100,000 rows typical)
  - Columns: AGREEMENTNO, MODEOFPAYMENT, AMOUNTPAID, RECEIPT ENTER DATE, Sub Zone, etc.
- **Sheet2**: Agreement master (AGREEMENTNO → Opening DPD, Slab)
  - 100,000+ rows (carries forward full portfolio)

### 2. **To_be_Disabled.xlsb**
- **CIF Level Disable**: CIF_NO → "Not Required" / "To be Disabled" / "CIF Level > 1.95 L"
- **Agreement Level Disble**: AGREEMENTNO → cash-mode status
- Used to flag compliance violations (disabled customers paying via cash/Airtel)

### 3. **CIFCL_CBSL_List.xlsx** (Employee Master)
- Must contain a **Mobile Number** column
- Used to flag receipts entered on agent phones vs. customer phones
- ~19,000 employee numbers typical

### 4. **DHC_Working_from_Last_Month.xlsx** (Carried-Forward Master)
- The **previous month's output** from this app (or the manual workbook)
- Specifically: **Look Up** sheet, columns J:N
  - AGREEMENTNO → CIF_NO, ZONE_NEW, SUB_REGION, OPENING_SLAB
  - ~5,200 agreements "known" from prior months
- **New agreements** (first appearance in this month's DCR) will show blank CIF — **you fill these in manually** and use the output as next month's input

---

## 📄 Output Sheets (5 Summary Tabs)

All sheets have **merged headers, subtotals, and professional formatting** matching the original DHC_Working.xlsx layout.

### **1. Receipt made summary**
- **Layout**: Updated/Pending vs Bounced/Cancelled (two side-by-side blocks)
- **Rows**: RECEIPT STATUS × PAYMENT MODE
- **Columns**: Status codes (Cleared, Deposit, Bounced, Cxn, Pending) + Grand Total
- **Values**: Counts (number of receipts)
- **Matches**: Your current manual pivot, exact row/column labels

### **2. RTGS Summary**
- **Left mini-block**: Online Payment sources (BBPS, CCP - QR, CCP - Bitly, etc.) with receipt counts
- **Main matrix**: 9 Sub Zones (EAST_1, NORTH_2, SOUTH_1, etc.) × 4 Receipt Types (OD, Settlement, Part Payment, Other OD) × 3 TAT buckets
- **TAT buckets**: < 4 Days, 5–10 Days, > 10 Days
- **Values**: Count + Value (Cr) for each cell

### **3. Cash Mode Validat Summary**
- **Rows**: CIF, Zone, Sub Region, Slab
- **Columns**: Receipt date (dynamic, one per day with violations)
- **Values**: Amount paid (cash/Airtel)
- **Compliance check**: Only shows agreements on the disable lists who paid via cash despite restrictions
- **Empty if clean**: Zero violations = good!

### **4. Delay in RCPTING Summary**
- **Full month** data (unlike RTGS which is filtered)
- Same structure as RTGS Summary: Zone × Receipt Type × TAT aging matrix
- Shows **aging** of all receipts (how old when entered into system)
- Helps identify bottlenecks (many > 10 days old = slow receipting)

### **5. RCPT CXN** (Cancelled Receipt Register)
- **Auto-populated columns**: Receipt No, Date, Amount, Receipt Type, Payment Mode, Zone, Agreement No, etc.
- **Remarks**: Left BLANK by design — you fill in why each receipt was cancelled
- **Counts**: Typically 10–50 cancelled receipts per month

---

## 🔧 Architecture & Logic

### Data Flow
```
DCR Sheet1 (receipts)
    ↓
[VLOOKUP replacements via pandas .merge()]
    ├─ → To_be_Disabled (compliance flags)
    ├─ → Employee mobiles (agent phone flag)
    ├─ → DCR Sheet2 (opening slab)
    └─ → Prior Look Up master (CIF/Zone carry-forward)
    ↓
DCR Tab (17,000 rows × 87 cols: raw + 11 derived)
    ↓
[Filter to RTGS mode, compute ageing]
    ↓
RTGS Tab (1,800 rows, subset)
    ↓
[5 aggregations: COUNTIFS/SUMIFS replicas via pandas pivot]
    ↓
5 Summary DataFrames → Excel with merged headers & formatting
```

### Key Differences from Manual Workbook
| Aspect | Manual | Automation |
|--------|--------|-----------|
| **Formula brittleness** | External links break; #REF! errors | Zero dependencies, fully self-contained |
| **Speed** | 20–30 min manual (if all data ready) | ~40 sec (no human input) |
| **Errors** | VLOOKUP typos, stale cache, forgot a formula | Deterministic Python — same result every run |
| **Maintenance** | Formula edits, copy-paste risks | Code in `etl.py`, version controlled, tested |

### The One Gap
**CIF/Zone/Sub Region mapping for NEW agreements** — ~5,200 agreements known from prior months, but if a customer's first receipt appears this month, they need manual CIF lookup from your LMS or a separate database. The app **flags these clearly** (see "New Agreements" metric on Streamlit dashboard) so you can fill them in once and use the output as next month's input.

---

## 💡 Usage Tips

### Month-to-Month Workflow
1. **Prepare files**: Get DCR, disable lists, employee master, and **last month's output** from the app
2. **Upload** all 4 files to the Streamlit app
3. **Click Process** — wait ~1 minute
4. **Preview** the 5 sheets (tabs at bottom of page)
5. **Download** the Excel output
6. **Next month**, use this month's output as your "Previous DHC Working" input

### New Agreements (Brand-New CIFCL Members)
- The "New Agreements" metric shows how many have no CIF mapping yet
- **Fill in manually** in the output's "Look Up Master" sheet (columns: CIF, ZONE_NEW, SUB_REGION)
- **Save the output** and use it as next month's "Previous DHC Working" input
- (This is the only manual step that can't be automated without direct LMS access)

### No More Manual RCPT CXN Remarks
- The Remarks column is intentionally blank
- **You** fill it in (e.g., "Duplicate receipt", "Customer requested", "Bank error")
- This is a judgment call that requires human insight, not automation

---

## 🐛 Troubleshooting

### "ERROR: Failed to load DCR"
- Confirm DCR is a `.xlsb` file with sheets named exactly `Sheet1` and `Sheet2`
- Check that Sheet1 has columns: AGREEMENTNO, MODEOFPAYMENT, RECEIPT ENTER DATE, etc.

### "0 RTGS Receipts in Output"
- Check your DCR — do receipts have Mode = 'RTGS'?
- The RTGS tab is **filtered**, so if all receipts are AIRTEL/CASH/ONLINE, RTGS will be empty (this is correct)

### "Many Agreements Marked 'New CIF'"
- These are receipts from customers not seen before in your organization
- Manually look them up in your LMS and fill in CIF/Zone/Sub Region in the Look Up Master sheet
- Use this output as next month's "Previous DHC Working" input to avoid re-doing them

### App Takes 2+ Minutes
- Large DCR files (50,000+ rows) can take longer
- This is normal; Python is optimized for the calculations
- If you routinely have 100,000+ receipts, consider splitting by date range and processing separately

---

## 📞 Support

For issues or feature requests:
1. Check the **Streamlit error message** (bottom of screen)
2. Try clearing browser cache & re-uploading files
3. Confirm file formats: `.xlsb` for DCR/disable, `.xlsx` for employee/prior working
4. See **ROADMAP.md** for full reverse-engineering notes and open questions

---

## 🎯 Next Steps After First Run

1. **Validate numbers** against your last month's manual workbook
   - Row counts, grand totals, zone breakdowns should match
   - If they don't, there's a bug — report it with a sample month's files

2. **Integrate downstream**
   - Export Streamlit output to your reporting tools, LMS, or email workflows
   - The Excel format is standard; no special post-processing needed

3. **Retire manual formulas**
   - Once you trust the automation (2–3 months of validation), stop maintaining the old DHC_Working.xlsx
   - Use the app output as your single source of truth

---

**Happy automating! 🚀**
