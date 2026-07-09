import pandas as pd
import numpy as np
import re
from collections import defaultdict
from rapidfuzz import fuzz
from parser import clean_invoice_number

def calculate_date_difference(date1, date2):
    """
    Calculates absolute difference in days between two dates.
    Returns np.nan if either date is invalid/NaT.
    """
    if pd.isna(date1) or pd.isna(date2):
        return np.nan
    return abs((date1 - date2).days)

def generate_gst_law_remark(row, source):
    """
    Generates standard GST compliance remarks based on GST rules & regulations.
    """
    if source == 'books_only':
        return "Only in Books. Document not uploaded by supplier. ITC cannot be claimed under Section 16(2)(aa) of CGST Act."
        
    # Extract variables from GSTR-2B side
    elg = row.get('gstr2b_itc_eligibility')
    rcm = row.get('gstr2b_rchrg')
    section = row.get('gstr2b_section')
    g3b = row.get('gstr2b_gstr3b_status')
    gstin = row.get('gstr2b_gstin')
    
    # 1. Blocked ITC Section 17(5)
    if elg == 'Ineligible':
        return "Blocked ITC under Section 17(5) of CGST Act (e.g. motor vehicles, catering, employee welfare). Claim not allowed."
        
    # 2. Reverse Charge (RCM)
    if rcm == 'Yes':
        return "Reverse Charge invoice (Section 9(3)/9(4) of CGST Act). Tax must be paid in cash by recipient. ITC claimable upon cash payment."
        
    # 3. ISD Invoices
    if section in ('ISD Invoices', 'ISD Amendments'):
        return "Distributed by Input Service Distributor (Section 20 of CGST Act). Claim allowed based on ISD invoice distribution details."
        
    # 4. Imports from ICEGATE
    if gstin == 'IMPORT':
        return "Import of goods from overseas. IGST paid under Customs Act (Section 3(7) of Customs Tariff Act). Subject to ICEGATE BOE matching."
        
    # 5. Imports from SEZ
    if section == 'Import from SEZ':
        return "Treated as Interstate supply from SEZ unit (Section 7(5)(b) of IGST Act). Subject to Bill of Entry validation."
        
    # 6. Supplier GSTR-3B status (Section 16(2)(aa))
    if g3b == 'No':
        return "Supplier GSTR-3B not filed. Claim is conditional on supplier filing GSTR-3B return under Section 16(2)(c)."

    return "Eligible Input Tax Credit (ITC) under Section 16 of CGST Act. Document uploaded and filed by supplier."

def get_numeric_suffix(clean_no):
    if not clean_no:
        return ""
    # Find trailing digits
    match = re.search(r'(\d+)$', clean_no)
    if match:
        return match.group(1)
    return ""

def consolidate_dataframe(df):
    """
    Groups multi-line invoice splits (e.g. by tax rate or cost center)
    by (supplier_gstin, doc_num, doc_type) and sums the monetary values.
    """
    if df is None or df.empty:
        return df
    
    # Ensure numeric columns are floats
    num_cols = ['taxable_val', 'igst', 'cgst', 'sgst', 'cess', 'total_val']
    df_copy = df.copy()
    for col in num_cols:
        if col in df_copy.columns:
            df_copy[col] = pd.to_numeric(df_copy[col], errors='coerce').fillna(0.0)
            
    # Group by key fields
    group_cols = ['supplier_gstin', 'doc_num', 'doc_type']
    
    # Check what columns actually exist in df
    agg_dict = {}
    for col in df_copy.columns:
        if col in num_cols:
            agg_dict[col] = 'sum'
        elif col not in group_cols:
            agg_dict[col] = 'first'
            
    df_consolidated = df_copy.groupby(group_cols, as_index=False).agg(agg_dict)
    
    # Round columns to 2 decimals
    for col in num_cols:
        if col in df_consolidated.columns:
            df_consolidated[col] = df_consolidated[col].round(2)
            
    return df_consolidated

def reconcile_data(books_df, gstr2b_df, val_tolerance=10.0, date_tolerance_days=7, fuzzy_threshold=85.0, tax_tolerance=10.0):
    """
    Executes an advanced 7-level matching strategy to reconcile Purchase Register and GSTR-2B.
    Groups multi-line items beforehand to resolve invoice discrepancies.
    """
    # 1. Consolidate multi-line splits
    b_df = consolidate_dataframe(books_df)
    g_df = consolidate_dataframe(gstr2b_df)
    
    # Add unique indexing columns for tracking matches
    b_df['books_idx'] = b_df.index
    g_df['gstr2b_idx'] = g_df.index

    # Convert to list of dicts for ultra-fast loop lookups
    b_rows = b_df.to_dict('records')
    g_rows = g_df.to_dict('records')

    matched_books = set()
    matched_gstr2b = set()
    reconciled_rows = []

    # Helper function to create combined record
    def create_reconciled_row(b_row, g_row, status, match_level):
        taxable_diff = round((b_row['taxable_val'] - g_row['taxable_val']), 2) if b_row is not None and g_row is not None else np.nan
        igst_diff = round((b_row['igst'] - g_row['igst']), 2) if b_row is not None and g_row is not None else np.nan
        cgst_diff = round((b_row['cgst'] - g_row['cgst']), 2) if b_row is not None and g_row is not None else np.nan
        sgst_diff = round((b_row['sgst'] - g_row['sgst']), 2) if b_row is not None and g_row is not None else np.nan
        cess_diff = round((b_row['cess'] - g_row['cess']), 2) if b_row is not None and g_row is not None else np.nan
        days_diff = calculate_date_difference(b_row['doc_date'], g_row['doc_date']) if b_row is not None and g_row is not None else np.nan

        # Generate descriptive remarks
        remarks = ""
        if status == 'Only in Books':
            remarks = "Only in Books. Document has not been uploaded by supplier to GST Portal. Hold ITC and follow up."
        elif status == 'Only in GSTR-2B':
            remarks = "Only in GSTR-2B. Document filed by supplier but entry is missing in your accounting books."
        elif b_row is not None and g_row is not None:
            # Reconciled record
            reasons = []
            if abs(taxable_diff) > val_tolerance:
                reasons.append(f"Taxable value diff Rs.{taxable_diff:,.2f} (Books: Rs.{b_row['taxable_val']:,.2f}, 2B: Rs.{g_row['taxable_val']:,.2f})")
            if not pd.isna(days_diff) and days_diff > date_tolerance_days:
                b_date = b_row['doc_date'].strftime('%d-%m-%Y') if not pd.isna(b_row['doc_date']) else "NaT"
                g_date = g_row['doc_date'].strftime('%d-%m-%Y') if not pd.isna(g_row['doc_date']) else "NaT"
                reasons.append(f"Date difference of {int(days_diff)} days (Books: {b_date}, GSTR-2B: {g_date})")
            if b_row['doc_num'] != g_row['doc_num']:
                reasons.append(f"Invoice number format variance (Books: '{b_row['doc_num']}', GSTR-2B: '{g_row['doc_num']}')")
            
            tax_mismatch_reasons = []
            if abs(igst_diff) > tax_tolerance:
                tax_mismatch_reasons.append(f"IGST diff: Rs.{igst_diff:,.2f}")
            if abs(cgst_diff) > tax_tolerance:
                tax_mismatch_reasons.append(f"CGST diff: Rs.{cgst_diff:,.2f}")
            if abs(sgst_diff) > tax_tolerance:
                tax_mismatch_reasons.append(f"SGST diff: Rs.{sgst_diff:,.2f}")
            if abs(cess_diff) > tax_tolerance:
                tax_mismatch_reasons.append(f"Cess diff: Rs.{cess_diff:,.2f}")
                
            if tax_mismatch_reasons:
                reasons.append("Tax values variance exceeds tolerance (" + ", ".join(tax_mismatch_reasons) + ")")
                status = "Tax Mismatch"
            elif abs(taxable_diff) > val_tolerance:
                status = "Value Mismatch"
            elif not pd.isna(days_diff) and days_diff > date_tolerance_days:
                status = "Date Mismatch"
                
            if reasons:
                remarks = f"Matched with warning ({match_level}): " + "; ".join(reasons)
            else:
                remarks = f"Exact match ({match_level}): Document keys and values reconcile perfectly."

        # Determine ITC action
        if status in ('Matched', 'Fuzzy Match', 'Matched (Amended)', 'Matched (Value-based)', 'Date Mismatch', 'Value Mismatch'):
            if g_row['itc_eligibility'] == 'Eligible':
                itc_action = 'ITC Claimable'
            else:
                itc_action = 'ITC Blocked (Ineligible)'
        elif 'Mismatch' in status or 'Discrepancy' in status or 'Tax Mismatch' in status:
            if g_row['itc_eligibility'] == 'Eligible':
                if abs(taxable_diff) <= val_tolerance:
                    itc_action = 'ITC Claimable'
                else:
                    if b_row['taxable_val'] > g_row['taxable_val']:
                        itc_action = 'Discrepancy (Claim GSTR-2B Value)'
                    else:
                        itc_action = 'Discrepancy (Claim Books Value)'
            else:
                itc_action = 'ITC Blocked (Ineligible)'
        elif status == 'Only in Books':
            itc_action = 'Pending supplier filing (Hold ITC)'
        elif status == 'Only in GSTR-2B':
            if g_row['itc_eligibility'] == 'Eligible':
                itc_action = 'Unrecorded in Books (Missing Entry)'
            else:
                itc_action = 'Unrecorded & Ineligible (Blocked)'
        else:
            itc_action = 'Review Required'

        row = {
            'reco_status': status,
            'match_level': match_level,
            'itc_action': itc_action,
            'remarks': remarks,
            
            # Books fields
            'books_gstin': b_row['supplier_gstin'] if b_row is not None else None,
            'books_supplier_name': b_row['supplier_name'] if b_row is not None else None,
            'books_doc_num': b_row['doc_num'] if b_row is not None else None,
            'books_doc_date': b_row['doc_date'] if b_row is not None else pd.NaT,
            'books_doc_type': b_row['doc_type'] if b_row is not None else None,
            'books_taxable_val': b_row['taxable_val'] if b_row is not None else 0.0,
            'books_igst': b_row['igst'] if b_row is not None else 0.0,
            'books_cgst': b_row['cgst'] if b_row is not None else 0.0,
            'books_sgst': b_row['sgst'] if b_row is not None else 0.0,
            'books_cess': b_row['cess'] if b_row is not None else 0.0,
            'books_total_val': b_row['total_val'] if b_row is not None else 0.0,
            'books_pos': b_row['pos'] if b_row is not None else "",
            'books_rchrg': b_row['rchrg'] if b_row is not None else "No",
            'books_pr_period': b_row['pr_period'] if b_row is not None else None,
            
            # GSTR-2B fields
            'gstr2b_gstin': g_row['supplier_gstin'] if g_row is not None else None,
            'gstr2b_supplier_name': g_row['supplier_name'] if g_row is not None else None,
            'gstr2b_doc_num': g_row['doc_num'] if g_row is not None else None,
            'gstr2b_doc_date': g_row['doc_date'] if g_row is not None else pd.NaT,
            'gstr2b_doc_type': g_row['doc_type'] if g_row is not None else None,
            'gstr2b_taxable_val': g_row['taxable_val'] if g_row is not None else 0.0,
            'gstr2b_igst': g_row['igst'] if g_row is not None else 0.0,
            'gstr2b_cgst': g_row['cgst'] if g_row is not None else 0.0,
            'gstr2b_sgst': g_row['sgst'] if g_row is not None else 0.0,
            'gstr2b_cess': g_row['cess'] if g_row is not None else 0.0,
            'gstr2b_total_val': g_row['total_val'] if g_row is not None else 0.0,
            'gstr2b_pos': g_row['pos'] if g_row is not None else "",
            'gstr2b_rchrg': g_row['rchrg'] if g_row is not None else "No",
            'gstr2b_itc_eligibility': g_row['itc_eligibility'] if g_row is not None else None,
            'gstr2b_filing_date': g_row['filing_date'] if g_row is not None else pd.NaT,
            'gstr2b_gstr3b_status': g_row['gstr3b_status'] if g_row is not None else "No",
            'gstr2b_section': g_row['section'] if g_row is not None else "",
            'gstr2b_rtn_period': g_row['rtn_period'] if g_row is not None else None,
            'gstr2b_source_file': g_row['source_file'] if g_row is not None else None,
            
            # Variances
            'taxable_val_diff': taxable_diff,
            'igst_diff': igst_diff,
            'cgst_diff': cgst_diff,
            'sgst_diff': sgst_diff,
            'cess_diff': cess_diff,
            'days_diff': days_diff
        }
        
        # Add law remarks
        if b_row is None and g_row is not None:
            row['gst_law_remark'] = generate_gst_law_remark(row, 'gstr2b_only')
        elif b_row is not None and g_row is None:
            row['gst_law_remark'] = generate_gst_law_remark(row, 'books_only')
        else:
            row['gst_law_remark'] = generate_gst_law_remark(row, 'matched')
            
        return row

    # Build Index maps for O(1) matching
    g_by_key = defaultdict(list)
    g_by_amended_key = defaultdict(list)
    g_by_suffix = defaultdict(list)
    g_by_value = defaultdict(list)
    
    for g_row in g_rows:
        gstin = g_row['supplier_gstin']
        doc_type = g_row['doc_type']
        clean_no = g_row['clean_doc_num']
        
        # Key: (GSTIN, Doc Type, Clean Invoice No)
        g_by_key[(gstin, doc_type, clean_no)].append(g_row)
        
        # Amendment
        if g_row.get('is_amended') and g_row.get('original_doc_num'):
            clean_original = clean_invoice_number(g_row['original_doc_num'])
            if clean_original:
                g_by_amended_key[(gstin, doc_type, clean_original)].append(g_row)
                
        # Suffix
        suffix = get_numeric_suffix(clean_no)
        if suffix:
            g_by_suffix[(gstin, doc_type, suffix)].append(g_row)
            
        # Value-based
        g_by_value[(gstin, doc_type, round(g_row['taxable_val'], 2))].append(g_row)

    # --- LEVEL 1 & 2: Exact Key Match (GSTIN + Clean Invoice No + Doc Type) ---
    for b_row in b_rows:
        b_idx = b_row['books_idx']
        gstin = b_row['supplier_gstin']
        doc_type = b_row['doc_type']
        clean_no = b_row['clean_doc_num']
        
        if not clean_no:
            continue
            
        candidates = g_by_key.get((gstin, doc_type, clean_no), [])
        available_candidates = [g for g in candidates if g['gstr2b_idx'] not in matched_gstr2b]
        
        if available_candidates:
            # Match candidate with minimum tax difference
            best_candidate = min(
                available_candidates,
                key=lambda g: (
                    abs(b_row['igst'] - g['igst']) +
                    abs(b_row['cgst'] - g['cgst']) +
                    abs(b_row['sgst'] - g['sgst']) +
                    abs(b_row['cess'] - g['cess'])
                )
            )
            
            g_idx = best_candidate['gstr2b_idx']
            matched_books.add(b_idx)
            matched_gstr2b.add(g_idx)
            
            # Check tolerances to separate Level 1 and Level 2
            taxable_diff = abs(b_row['taxable_val'] - best_candidate['taxable_val'])
            igst_diff = abs(b_row['igst'] - best_candidate['igst'])
            cgst_diff = abs(b_row['cgst'] - best_candidate['cgst'])
            sgst_diff = abs(b_row['sgst'] - best_candidate['sgst'])
            cess_diff = abs(b_row['cess'] - best_candidate['cess'])
            days_diff = calculate_date_difference(b_row['doc_date'], best_candidate['doc_date'])
            
            is_within_tol = (
                taxable_diff <= val_tolerance and
                igst_diff <= tax_tolerance and
                cgst_diff <= tax_tolerance and
                sgst_diff <= tax_tolerance and
                cess_diff <= tax_tolerance and
                (pd.isna(days_diff) or days_diff <= date_tolerance_days)
            )
            
            if is_within_tol:
                status = 'Matched'
                match_level = 'Level 1: Exact Match'
            else:
                status = 'Tax Mismatch' if (igst_diff > tax_tolerance or cgst_diff > tax_tolerance or sgst_diff > tax_tolerance or cess_diff > tax_tolerance) else ('Value Mismatch' if taxable_diff > val_tolerance else 'Date Mismatch')
                match_level = 'Level 2: Exact Key Match with Mismatch'
                
            reconciled_rows.append(create_reconciled_row(b_row, best_candidate, status, match_level))

    # --- LEVEL 5: Amendment Matching (Original Doc Number lookup) ---
    for b_row in b_rows:
        b_idx = b_row['books_idx']
        if b_idx in matched_books:
            continue
            
        gstin = b_row['supplier_gstin']
        doc_type = b_row['doc_type']
        clean_no = b_row['clean_doc_num']
        
        if not clean_no:
            continue
            
        candidates = g_by_amended_key.get((gstin, doc_type, clean_no), [])
        available_candidates = [g for g in candidates if g['gstr2b_idx'] not in matched_gstr2b]
        
        if available_candidates:
            best_candidate = available_candidates[0]
            g_idx = best_candidate['gstr2b_idx']
            matched_books.add(b_idx)
            matched_gstr2b.add(g_idx)
            
            taxable_diff = abs(b_row['taxable_val'] - best_candidate['taxable_val'])
            igst_diff = abs(b_row['igst'] - best_candidate['igst'])
            cgst_diff = abs(b_row['cgst'] - best_candidate['cgst'])
            sgst_diff = abs(b_row['sgst'] - best_candidate['sgst'])
            cess_diff = abs(b_row['cess'] - best_candidate['cess'])
            
            is_within_tol = (
                taxable_diff <= val_tolerance and
                igst_diff <= tax_tolerance and
                cgst_diff <= tax_tolerance and
                sgst_diff <= tax_tolerance and
                cess_diff <= tax_tolerance
            )
            
            if is_within_tol:
                status = 'Matched (Amended)'
            else:
                status = 'Tax Mismatch (Amended)' if (igst_diff > tax_tolerance or cgst_diff > tax_tolerance or sgst_diff > tax_tolerance or cess_diff > tax_tolerance) else 'Value Mismatch (Amended)'
                
            reconciled_rows.append(create_reconciled_row(b_row, best_candidate, status, 'Level 5: Amendment Match'))

    # --- LEVEL 3 & 4: Suffix/Numeric Match ---
    for b_row in b_rows:
        b_idx = b_row['books_idx']
        if b_idx in matched_books:
            continue
            
        gstin = b_row['supplier_gstin']
        doc_type = b_row['doc_type']
        clean_no = b_row['clean_doc_num']
        
        suffix = get_numeric_suffix(clean_no)
        if not suffix:
            continue
            
        candidates = g_by_suffix.get((gstin, doc_type, suffix), [])
        available_candidates = [g for g in candidates if g['gstr2b_idx'] not in matched_gstr2b]
        
        if available_candidates:
            # Check short suffix constraint to prevent false matches
            if len(suffix) < 3:
                # Require values to match closely
                available_candidates = [
                    g for g in available_candidates
                    if abs(b_row['taxable_val'] - g['taxable_val']) <= val_tolerance
                ]
                if not available_candidates:
                    continue
            
            best_candidate = min(
                available_candidates,
                key=lambda g: (
                    abs(b_row['igst'] - g['igst']) +
                    abs(b_row['cgst'] - g['cgst']) +
                    abs(b_row['sgst'] - g['sgst']) +
                    abs(b_row['cess'] - g['cess'])
                )
            )
            
            g_idx = best_candidate['gstr2b_idx']
            matched_books.add(b_idx)
            matched_gstr2b.add(g_idx)
            
            taxable_diff = abs(b_row['taxable_val'] - best_candidate['taxable_val'])
            igst_diff = abs(b_row['igst'] - best_candidate['igst'])
            cgst_diff = abs(b_row['cgst'] - best_candidate['cgst'])
            sgst_diff = abs(b_row['sgst'] - best_candidate['sgst'])
            cess_diff = abs(b_row['cess'] - best_candidate['cess'])
            days_diff = calculate_date_difference(b_row['doc_date'], best_candidate['doc_date'])
            
            is_within_tol = (
                taxable_diff <= val_tolerance and
                igst_diff <= tax_tolerance and
                cgst_diff <= tax_tolerance and
                sgst_diff <= tax_tolerance and
                cess_diff <= tax_tolerance and
                (pd.isna(days_diff) or days_diff <= date_tolerance_days)
            )
            
            if is_within_tol:
                status = 'Matched'
                match_level = 'Level 3: Suffix/Numeric Match'
            else:
                status = 'Tax Mismatch' if (igst_diff > tax_tolerance or cgst_diff > tax_tolerance or sgst_diff > tax_tolerance or cess_diff > tax_tolerance) else ('Value Mismatch' if taxable_diff > val_tolerance else 'Date Mismatch')
                match_level = 'Level 4: Suffix/Numeric Match with Mismatch'
                
            reconciled_rows.append(create_reconciled_row(b_row, best_candidate, status, match_level))

    # --- LEVEL 6: Fuzzy Invoice Number Match within Supplier GSTIN ---
    unmatched_books_list = [b for b in b_rows if b['books_idx'] not in matched_books]
    unmatched_gstr2b_list = [g for g in g_rows if g['gstr2b_idx'] not in matched_gstr2b]
    
    unmatched_b_by_gstin = defaultdict(list)
    for b_row in unmatched_books_list:
        unmatched_b_by_gstin[b_row['supplier_gstin']].append(b_row)
        
    unmatched_g_by_gstin = defaultdict(list)
    for g_row in unmatched_gstr2b_list:
        unmatched_g_by_gstin[g_row['supplier_gstin']].append(g_row)
        
    unique_gstins = set(unmatched_b_by_gstin.keys()).intersection(set(unmatched_g_by_gstin.keys()))
    
    for gstin in unique_gstins:
        if gstin in ('IMPORT', 'UNKNOWN'):
            continue
            
        b_sub = unmatched_b_by_gstin[gstin]
        g_sub = unmatched_g_by_gstin[gstin]
        
        complexity = len(b_sub) * len(g_sub)
        if complexity > 100000:
            continue
            
        for b_row in b_sub:
            b_idx = b_row['books_idx']
            if b_idx in matched_books:
                continue
                
            best_score = 0
            best_candidate = None
            
            for g_row in g_sub:
                g_idx = g_row['gstr2b_idx']
                if g_idx in matched_gstr2b:
                    continue
                if b_row['doc_type'] != g_row['doc_type']:
                    continue
                    
                # Taxable value ratio check
                b_val = abs(b_row['taxable_val'])
                g_val = abs(g_row['taxable_val'])
                if b_val > 500 and g_val > 500:
                    val_ratio = min(b_val, g_val) / max(b_val, g_val)
                    if val_ratio < 0.75:
                        continue
                
                score = fuzz.ratio(b_row['doc_num'], g_row['doc_num'])
                clean_score = fuzz.ratio(b_row['clean_doc_num'], g_row['clean_doc_num'])
                max_score = max(score, clean_score)
                
                if max_score > best_score:
                    best_score = max_score
                    best_candidate = g_row
                    
            if best_score >= fuzzy_threshold and best_candidate is not None:
                g_idx = best_candidate['gstr2b_idx']
                matched_books.add(b_idx)
                matched_gstr2b.add(g_idx)
                
                taxable_diff = abs(b_row['taxable_val'] - best_candidate['taxable_val'])
                igst_diff = abs(b_row['igst'] - best_candidate['igst'])
                cgst_diff = abs(b_row['cgst'] - best_candidate['cgst'])
                sgst_diff = abs(b_row['sgst'] - best_candidate['sgst'])
                cess_diff = abs(b_row['cess'] - best_candidate['cess'])
                
                is_within_tol = (
                    taxable_diff <= val_tolerance and
                    igst_diff <= tax_tolerance and
                    cgst_diff <= tax_tolerance and
                    sgst_diff <= tax_tolerance and
                    cess_diff <= tax_tolerance
                )
                
                if is_within_tol:
                    status = 'Fuzzy Match'
                else:
                    status = 'Tax Mismatch (Fuzzy)' if (igst_diff > tax_tolerance or cgst_diff > tax_tolerance or sgst_diff > tax_tolerance or cess_diff > tax_tolerance) else 'Value Mismatch (Fuzzy)'
                    
                reconciled_rows.append(create_reconciled_row(
                    b_row, best_candidate, status, f'Level 6: Fuzzy Match (Score: {int(best_score)}%)'
                ))

    # --- LEVEL 7: Value-based Match (GSTIN + Doc Type + Value exact match) ---
    for b_row in b_rows:
        b_idx = b_row['books_idx']
        if b_idx in matched_books:
            continue
            
        gstin = b_row['supplier_gstin']
        doc_type = b_row['doc_type']
        taxable_val = round(b_row['taxable_val'], 2)
        
        candidates = g_by_value.get((gstin, doc_type, taxable_val), [])
        available_candidates = [g for g in candidates if g['gstr2b_idx'] not in matched_gstr2b]
        
        if available_candidates:
            # Check tight tax amount tolerance (allow within 2.0 to be extremely accurate)
            matching_value_candidate = None
            for g in available_candidates:
                tax_diff = (
                    abs(b_row['igst'] - g['igst']) +
                    abs(b_row['cgst'] - g['cgst']) +
                    abs(b_row['sgst'] - g['sgst']) +
                    abs(b_row['cess'] - g['cess'])
                )
                if tax_diff <= 2.0:
                    matching_value_candidate = g
                    break
            
            if matching_value_candidate is not None:
                g_idx = matching_value_candidate['gstr2b_idx']
                matched_books.add(b_idx)
                matched_gstr2b.add(g_idx)
                
                reconciled_rows.append(create_reconciled_row(
                    b_row, matching_value_candidate, 'Matched (Value-based)', 'Level 7: Value-based Match'
                ))

    # --- ONLY IN BOOKS (Unmatched books entries) ---
    for b_row in b_rows:
        b_idx = b_row['books_idx']
        if b_idx not in matched_books:
            reconciled_rows.append(create_reconciled_row(b_row, None, 'Only in Books', 'Unmatched'))

    # --- ONLY IN GSTR-2B (Unmatched portal entries) ---
    for g_row in g_rows:
        g_idx = g_row['gstr2b_idx']
        if g_idx not in matched_gstr2b:
            reconciled_rows.append(create_reconciled_row(None, g_row, 'Only in GSTR-2B', 'Unmatched'))

    if not reconciled_rows:
        return pd.DataFrame()
        
    df_reco = pd.DataFrame(reconciled_rows)
    
    # Clean temporary indexes
    if 'books_idx' in df_reco.columns:
        df_reco = df_reco.drop(columns=['books_idx'], errors='ignore')
    if 'gstr2b_idx' in df_reco.columns:
        df_reco = df_reco.drop(columns=['gstr2b_idx'], errors='ignore')
        
    return df_reco

def generate_supplier_summary(df_reco):
    """
    Groups reconciled data by Supplier GSTIN to summarize alignment performance.
    Handles advanced matching statuses robustly.
    """
    if df_reco.empty:
        return pd.DataFrame()
        
    # Standardize GSTIN identifier: prefer Books GSTIN, fallback to GSTR-2B GSTIN
    df_reco['summary_gstin'] = df_reco['books_gstin'].fillna(df_reco['gstr2b_gstin'])
    df_reco['summary_supplier_name'] = df_reco['books_supplier_name'].fillna(df_reco['gstr2b_supplier_name'])
    
    summary_data = []
    
    for gstin, group in df_reco.groupby('summary_gstin'):
        supplier_name = group['summary_supplier_name'].dropna().iloc[0] if not group['summary_supplier_name'].dropna().empty else "Unknown Supplier"
        
        total_books_invoices = int(group['books_doc_num'].dropna().nunique())
        total_gstr2b_invoices = int(group['gstr2b_doc_num'].dropna().nunique())
        
        books_taxable = round(group['books_taxable_val'].sum(), 2)
        books_igst = round(group['books_igst'].sum(), 2)
        books_cgst = round(group['books_cgst'].sum(), 2)
        books_sgst = round(group['books_sgst'].sum(), 2)
        books_total_itc = round(books_igst + books_cgst + books_sgst, 2)
        
        gstr2b_taxable = round(group['gstr2b_taxable_val'].sum(), 2)
        gstr2b_igst = round(group['gstr2b_igst'].sum(), 2)
        gstr2b_cgst = round(group['gstr2b_cgst'].sum(), 2)
        gstr2b_sgst = round(group['gstr2b_sgst'].sum(), 2)
        gstr2b_total_itc = round(gstr2b_igst + gstr2b_cgst + gstr2b_sgst, 2)
        
        taxable_diff = round(books_taxable - gstr2b_taxable, 2)
        itc_diff = round(books_total_itc - gstr2b_total_itc, 2)
        
        # Match count calculations
        # Any status starting with 'Matched' (like Matched, Matched (Amended), Matched (Value-based)) or 'Fuzzy Match' and not containing 'Mismatch'
        is_matched_mask = group['reco_status'].str.contains('Matched|Fuzzy', case=False, na=False) & ~group['reco_status'].str.contains('Mismatch', case=False, na=False)
        matched_docs_count = int(group[is_matched_mask].shape[0])
        
        # Detailed counts for backward compatibility with UI expectation
        matched_count = int(group[group['reco_status'] == 'Matched'].shape[0])
        fuzzy_count = int(group[group['reco_status'].isin(['Fuzzy Match', 'Matched (Value-based)'])].shape[0])
        amended_count = int(group[group['reco_status'].isin(['Matched (Amended)', 'Tax Mismatch (Amended)', 'Value Mismatch (Amended)'])].shape[0])
        mismatch_count = int(group[group['reco_status'].str.contains('Mismatch', na=False)].shape[0])
        only_books_count = int(group[group['reco_status'] == 'Only in Books'].shape[0])
        only_gstr2b_count = int(group[group['reco_status'] == 'Only in GSTR-2B'].shape[0])
        
        total_docs = len(group)
        match_rate = round((matched_docs_count / total_docs * 100.0), 1) if total_docs > 0 else 0.0
        
        summary_data.append({
            'supplier_gstin': gstin,
            'supplier_name': supplier_name,
            'books_invoice_count': total_books_invoices,
            'gstr2b_invoice_count': total_gstr2b_invoices,
            'books_taxable_val': books_taxable,
            'books_total_itc': books_total_itc,
            'gstr2b_taxable_val': gstr2b_taxable,
            'gstr2b_total_itc': gstr2b_total_itc,
            'taxable_val_diff': taxable_diff,
            'itc_diff': itc_diff,
            'match_rate_pct': match_rate,
            'matched_count': matched_count,
            'fuzzy_count': fuzzy_count,
            'mismatch_count': mismatch_count,
            'only_books_count': only_books_count,
            'only_gstr2b_count': only_gstr2b_count
        })
        
    return pd.DataFrame(summary_data)
