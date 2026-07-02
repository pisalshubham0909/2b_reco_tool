import pandas as pd
import numpy as np
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

def reconcile_data(books_df, gstr2b_df, val_tolerance=10.0, date_tolerance_days=7, fuzzy_threshold=85.0):
    """
    Executes a 4-level matching strategy to reconcile Purchase Register and GSTR-2B.
    Optimized to complete in linear O(N) time using hash-map indices.
    """
    # 1. Prepare dataframes
    b_df = books_df.copy()
    g_df = gstr2b_df.copy()
    
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
        days_diff = calculate_date_difference(b_row['doc_date'], g_row['doc_date']) if b_row is not None and g_row is not None else np.nan

        # Generate descriptive remarks
        remarks = ""
        if status == 'Matched':
            remarks = "Exact match: Document keys and values reconcile perfectly."
        elif status == 'Matched (Amended)':
            remarks = f"Matched via amended document link (Original doc number: {g_row['original_doc_num']})."
        elif status == 'Fuzzy Match':
            score_str = match_level.split("Score: ")[-1].replace(")", "") if "Score: " in match_level else "85"
            remarks = f"Fuzzy match on doc number (Similarity score: {score_str}%). Minor prefix/formatting difference."
        elif status == 'Only in Books':
            remarks = "Only in Books. Document has not been uploaded by supplier to GST Portal. Hold ITC and follow up."
        elif status == 'Only in GSTR-2B':
            remarks = "Only in GSTR-2B. Document filed by supplier but entry is missing in your accounting books."
        else:
            # This is a mismatch
            reasons = []
            if abs(taxable_diff) > val_tolerance:
                reasons.append(f"Taxable value diff ₹{taxable_diff:,.2f} (Books: ₹{b_row['taxable_val']:,.2f}, 2B: ₹{g_row['taxable_val']:,.2f})")
            if not pd.isna(days_diff) and days_diff > date_tolerance_days:
                b_date = b_row['doc_date'].strftime('%d-%m-%Y') if not pd.isna(b_row['doc_date']) else "NaT"
                g_date = g_row['doc_date'].strftime('%d-%m-%Y') if not pd.isna(g_row['doc_date']) else "NaT"
                reasons.append(f"Date difference of {int(days_diff)} days (Books: {b_date}, GSTR-2B: {g_date})")
            
            tax_mismatch_reasons = []
            if abs(igst_diff) > 1.0:
                tax_mismatch_reasons.append(f"IGST diff: ₹{igst_diff:,.2f}")
            if abs(cgst_diff) > 1.0:
                tax_mismatch_reasons.append(f"CGST diff: ₹{cgst_diff:,.2f}")
            if abs(sgst_diff) > 1.0:
                tax_mismatch_reasons.append(f"SGST diff: ₹{sgst_diff:,.2f}")
                
            if tax_mismatch_reasons:
                reasons.append("Tax values variance (" + ", ".join(tax_mismatch_reasons) + ")")
                
            remarks = "Discrepancy: " + "; ".join(reasons)

        # Determine ITC action
        if status in ('Matched', 'Fuzzy Match', 'Matched (Amended)'):
            if g_row['itc_eligibility'] == 'Eligible':
                itc_action = 'ITC Claimable'
            else:
                itc_action = 'ITC Blocked (Ineligible)'
        elif 'Mismatch' in status or 'Discrepancy' in status:
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
            'books_section': b_row['section'] if b_row is not None else "",
            
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
    from collections import defaultdict
    g_by_key = defaultdict(list)
    g_by_amended_key = defaultdict(list)
    
    for g_row in g_rows:
        key = (g_row['supplier_gstin'], g_row['doc_type'], g_row['clean_doc_num'])
        g_by_key[key].append(g_row)
        
        if g_row.get('is_amended') and g_row.get('original_doc_num'):
            clean_original = clean_invoice_number(g_row['original_doc_num'])
            if clean_original:
                am_key = (g_row['supplier_gstin'], g_row['doc_type'], clean_original)
                g_by_amended_key[am_key].append(g_row)

    # --- LEVEL 1: Exact Match (GSTIN + Clean Invoice No + Doc Type + Within tolerance) ---
    for b_row in b_rows:
        b_idx = b_row['books_idx']
        gstin = b_row['supplier_gstin']
        doc_type = b_row['doc_type']
        clean_no = b_row['clean_doc_num']
        
        if not clean_no:
            continue
            
        key = (gstin, doc_type, clean_no)
        candidates = g_by_key.get(key, [])
        
        best_candidate = None
        min_val_diff = float('inf')
        
        for g_row in candidates:
            g_idx = g_row['gstr2b_idx']
            if g_idx in matched_gstr2b:
                continue
                
            val_diff = abs(b_row['taxable_val'] - g_row['taxable_val'])
            days_diff = calculate_date_difference(b_row['doc_date'], g_row['doc_date'])
            
            if val_diff <= val_tolerance and (pd.isna(days_diff) or days_diff <= date_tolerance_days):
                if val_diff < min_val_diff:
                    min_val_diff = val_diff
                    best_candidate = g_row
                    
        if best_candidate is not None:
            g_idx = best_candidate['gstr2b_idx']
            matched_books.add(b_idx)
            matched_gstr2b.add(g_idx)
            reconciled_rows.append(create_reconciled_row(b_row, best_candidate, 'Matched', 'Level 1: Exact Match'))

    # --- LEVEL 2: Exact Key Match with Date/Value Discrepancy ---
    for b_row in b_rows:
        b_idx = b_row['books_idx']
        if b_idx in matched_books:
            continue
            
        gstin = b_row['supplier_gstin']
        doc_type = b_row['doc_type']
        clean_no = b_row['clean_doc_num']
        
        if not clean_no:
            continue
            
        key = (gstin, doc_type, clean_no)
        candidates = g_by_key.get(key, [])
        
        available_candidates = [g for g in candidates if g['gstr2b_idx'] not in matched_gstr2b]
        
        if available_candidates:
            best_candidate = None
            min_val_diff = float('inf')
            for g_row in available_candidates:
                val_diff = abs(b_row['taxable_val'] - g_row['taxable_val'])
                if val_diff < min_val_diff:
                    min_val_diff = val_diff
                    best_candidate = g_row
            
            if best_candidate is not None:
                g_idx = best_candidate['gstr2b_idx']
                matched_books.add(b_idx)
                matched_gstr2b.add(g_idx)
                
                val_diff = abs(b_row['taxable_val'] - best_candidate['taxable_val'])
                days_diff = calculate_date_difference(b_row['doc_date'], best_candidate['doc_date'])
                
                # Determine mismatch nature
                mismatch_reasons = []
                if val_diff > val_tolerance:
                    mismatch_reasons.append('Value Mismatch')
                if not pd.isna(days_diff) and days_diff > date_tolerance_days:
                    mismatch_reasons.append('Date Mismatch')
                
                status = " & ".join(mismatch_reasons) if mismatch_reasons else 'Matched'
                reconciled_rows.append(create_reconciled_row(b_row, best_candidate, status, 'Level 2: ID Match with Discrepancy'))

    # --- LEVEL 3: Amendment Matching (Original Doc Number lookup) ---
    for b_row in b_rows:
        b_idx = b_row['books_idx']
        if b_idx in matched_books:
            continue
            
        gstin = b_row['supplier_gstin']
        doc_type = b_row['doc_type']
        clean_no = b_row['clean_doc_num']
        
        if not clean_no:
            continue
            
        key = (gstin, doc_type, clean_no)
        candidates = g_by_amended_key.get(key, [])
        
        available_candidates = [g for g in candidates if g['gstr2b_idx'] not in matched_gstr2b]
        
        if available_candidates:
            best_candidate = available_candidates[0]
            g_idx = best_candidate['gstr2b_idx']
            matched_books.add(b_idx)
            matched_gstr2b.add(g_idx)
            
            val_diff = abs(b_row['taxable_val'] - best_candidate['taxable_val'])
            status = 'Matched (Amended)' if val_diff <= val_tolerance else 'Value Mismatch (Amended)'
            reconciled_rows.append(create_reconciled_row(b_row, best_candidate, status, 'Level 3: Amendment Match'))

    # --- LEVEL 4: Fuzzy Invoice Number Match within Supplier GSTIN ---
    # Group unmatched records by GSTIN
    unmatched_books_list = [b for b in b_rows if b['books_idx'] not in matched_books]
    unmatched_gstr2b_list = [g for g in g_rows if g['gstr2b_idx'] not in matched_gstr2b]
    
    # Organize by GSTIN
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
        
        # O(N*M) Guard: check total pairwise comparison complexity for this supplier
        complexity = len(b_sub) * len(g_sub)
        
        if complexity > 100000:
            # Too complex for this supplier, skip fuzzy match to protect server responsiveness
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
                    
                # Optimization: Block comparison if taxable values are completely divergent (e.g. >25% diff)
                b_val = abs(b_row['taxable_val'])
                g_val = abs(g_row['taxable_val'])
                if b_val > 500 and g_val > 500:
                    val_ratio = min(b_val, g_val) / max(b_val, g_val)
                    if val_ratio < 0.75:
                        continue
                
                # Perform fuzzy similarity matching
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
                reconciled_rows.append(create_reconciled_row(
                    b_row, best_candidate, 'Fuzzy Match', f'Level 4: Fuzzy Match (Score: {int(best_score)}%)'
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
        matched_count = int(group[group['reco_status'] == 'Matched'].shape[0])
        fuzzy_count = int(group[group['reco_status'] == 'Fuzzy Match'].shape[0])
        amended_count = int(group[group['reco_status'].isin(['Matched (Amended)', 'Value Mismatch (Amended)'])].shape[0])
        mismatch_count = int(group[group['reco_status'].str.contains('Mismatch', na=False)].shape[0])
        only_books_count = int(group[group['reco_status'] == 'Only in Books'].shape[0])
        only_gstr2b_count = int(group[group['reco_status'] == 'Only in GSTR-2B'].shape[0])
        
        total_docs = len(group)
        match_rate = round(((matched_count + fuzzy_count + amended_count) / total_docs * 100.0), 1) if total_docs > 0 else 0.0
        
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
