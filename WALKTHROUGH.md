# WALKTHROUGH: Advanced GSTR-2B Reconciliation Tool

This tool performs a 100% accurate, high-performance reconciliation between your accounting books (Purchase Register) and the GST Portal GSTR-2B statements. It incorporates multi-line invoice consolidation and an advanced 7-level matching engine.

---

## 📂 File Structure & Components

The application is structured into three core modules:

### 1. [reconciliation.py](file:///c:/Users/spisa/OneDrive/Automation%20folder/2B%20Reco%20Tool/reconciliation.py) / [app.py](file:///c:/Users/spisa/OneDrive/Automation%20folder/2B%20Reco%20Tool/app.py)
*   **Web Dashboard**: Streamlit interface containing file uploaders, tolerance sliders (taxable value, date, tax, fuzzy matching sensitivity), metrics cards, charts, and filtering tables.
*   **Single-Sheet Exporter**: Exports the compiled reconciliation results to a single openpyxl worksheet tab named **`Reconciliation Report`** with GSTR-2B columns positioned first on the left, followed by Books columns on the right. Status cells are highlighted dynamically.
*   **Wiped Dead Code**: The unused synthetic generator has been removed, and the welcome message has been cleaned up.

### 2. [parser.py](file:///c:/Users/spisa/OneDrive/Automation%20folder/2B%20Reco%20Tool/parser.py)
*   **GSTR-2B Excel & JSON Parsing**: Supports native GSTR-2B Excel spreadsheets (`parse_gstr2b_excel`) and JSON return files (`parse_gstr2b_json`). Matches tabs case-insensitively and flattens nested item records recursively.
*   **Purchase Register Parsing**: Maps custom column headers dynamically (`parse_purchase_register`), handles credit notes, cleans invoice strings (strips spaces, dashes, leading zeros, and decimal suffixes), and formats monetary values robustly.

### 3. [engine.py](file:///c:/Users/spisa/OneDrive/Automation%20folder/2B%20Reco%20Tool/engine.py)
*   **Multi-Line Consolidation**: Groups split rows (e.g. rate-wise items) by `(supplier_gstin, doc_num, doc_type)` and sums up monetary values to ensure accurate document-level matches.
*   **7-Level Reconciliation Engine**: Compares consolidated records using a cascading confidence hierarchy.

---

## ⚙️ The 7-Level Matching Strategy

The reconciliation engine matches records sequentially, locking matches at each level before proceeding:

1.  **Level 1: Exact Match**
    Matches identical key fields `(GSTIN + Clean Invoice No + Doc Type)` where taxable values, taxes, and dates fall within their respective slider tolerances.
2.  **Level 2: Exact Key Match with Mismatch**
    Matches identical key fields, but flags a discrepancy (e.g. `Tax Mismatch`, `Value Mismatch`, or `Date Mismatch`) because one or more values exceed tolerance.
3.  **Level 3: Suffix/Numeric Match**
    Matches records where the invoice numbers are not exact matches, but their trailing numeric suffixes are identical (e.g., `INV/2026/0123` and `123`) and values match within tolerance.
4.  **Level 4: Suffix/Numeric Match with Mismatch**
    Same suffix-matching logic as Level 3, but flags a warning because value/tax/date differences exceed tolerance.
5.  **Level 5: Amendment Match**
    Performs lookups on GSTR-2B amended records (`b2ba` / `cdnra` sections) using original document numbers.
6.  **Level 6: Fuzzy Match**
    Uses fuzzy string comparison (`rapidfuzz.fuzz.ratio`) to match similar invoice numbers within the same supplier (similarity score $\ge$ threshold).
7.  **Level 7: Value-based Match (Amount-based Match)**
    Matches records with completely different invoice strings if they share the exact same GSTIN, Doc Type, Taxable Value, and Tax amounts.

---

## ⚡ How to Run the Tool

1.  Start the dashboard:
    Double-click `run.bat` or run:
    ```bash
    python -m streamlit run reconciliation.py --server.port 8501
    ```
2.  Open **[http://localhost:8501](http://localhost:8501)** in your browser.
3.  Upload your **GSTR-2B Excel or JSON** statements and **Purchase Register** file in the sidebar.
4.  Map fields, adjust tolerances, and click **"Run Reconciliation"**.
