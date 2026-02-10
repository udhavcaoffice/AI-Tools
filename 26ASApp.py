import streamlit as st
import pandas as pd
import io
import re
from pdf2image import convert_from_bytes
import pytesseract
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from fuzzywuzzy import fuzz

# --- Page Configuration ---
st.set_page_config(page_title="CA Udhav AI Tools", layout="wide")

st.title("CA Udhav Agarwalla - AI Utilities")
st.markdown("Select a tool below to process your audit files.")

# --- Tabs for the 3 Tools ---
tab1, tab2, tab3 = st.tabs(["📄 1. 26AS PDF to Excel", "📒 2. Tally to Summary", "🔍 3. Reconciliation"])

# ==========================================
# TOOL 1: 26AS PDF TO EXCEL (OCR)
# ==========================================
with tab1:
    st.header("Convert 26AS PDF to Excel")
    st.info("Note: This tool uses OCR. Large files may take a moment to process.")
    
    uploaded_pdf = st.file_uploader("Upload 26AS PDF", type="pdf", key="t1")
    
    if uploaded_pdf:
        if st.button("Convert to Excel", key="btn1"):
            with st.spinner("Scanning PDF... (this may take 30-60 seconds)"):
                try:
                    # Convert PDF bytes to images
                    images = convert_from_bytes(uploaded_pdf.read())
                    full_text = ""
                    
                    # Extract text from images
                    progress_bar = st.progress(0)
                    for i, image in enumerate(images):
                        text = pytesseract.image_to_string(image, config='--psm 4')
                        full_text += text + "\n"
                        progress_bar.progress((i + 1) / len(images))
                    
                    # Parse Text Logic (Your original logic)
                    data = []
                    lines = full_text.split('\n')
                    tan_loose_pattern = re.compile(r'[A-Z]{4}[0-9OIl]{5}[A-Z]')

                    for line in lines:
                        line = line.strip()
                        if len(line) < 15: continue
                        match = tan_loose_pattern.search(line)
                        if match:
                            tan_code = match.group()
                            clean_line = re.sub(r'\s+', ' ', line)
                            parts = clean_line.split(' ')
                            tan_idx = -1
                            for idx, part in enumerate(parts):
                                if tan_code in part:
                                    tan_idx = idx
                                    break
                            if tan_idx != -1:
                                name_parts = []
                                for j in range(tan_idx):
                                    w = parts[j]
                                    if len(w) > 1 and not re.match(r'^\d+$', w) and w.lower() not in ['sr', 'no']:
                                        name_parts.append(w)
                                party_name = " ".join(name_parts)
                                party_name = re.sub(r'^[^A-Z]+', '', party_name)

                                amounts = []
                                for token in parts[tan_idx+1:]:
                                    token_fix = token.replace('O','0').replace('o','0').replace('l','1').replace('I','1').replace('S','5')
                                    token_clean = re.sub(r'[^\d\.]', '', token_fix)
                                    if re.match(r'^\d+\.?\d{0,2}$', token_clean):
                                        try:
                                            val = float(token_clean)
                                            if val > 10:
                                                amounts.append(val)
                                        except: pass
                                
                                final_tax = 0.0
                                if len(amounts) >= 3:
                                    final_tax = amounts[1]
                                elif len(amounts) == 2:
                                    final_tax = amounts[1]
                                elif len(amounts) == 1:
                                    final_tax = amounts[0]
                                
                                if len(party_name) > 3 and final_tax > 0:
                                    data.append({
                                        'Name of Party': party_name,
                                        'Amount showing in 26AS': final_tax
                                    })

                    df = pd.DataFrame(data)
                    
                    if not df.empty:
                        # Clean and Deduplicate
                        df = df.drop_duplicates(subset=['Name of Party'], keep='first')
                        df = df.sort_values('Name of Party').reset_index(drop=True)
                        
                        # Generate Excel in Memory
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False, sheet_name='26AS Data')
                            ws = writer.sheets['26AS Data']
                            ws.column_dimensions['A'].width = 50
                            ws.column_dimensions['B'].width = 18
                        
                        output.seek(0)
                        
                        st.success(f"Success! Extracted {len(df)} rows.")
                        st.dataframe(df)
                        st.download_button(
                            label="Download 26AS Excel",
                            data=output,
                            file_name="26AS_Extracted_Data.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    else:
                        st.error("No valid data found in the PDF. Please ensure the scan quality is good.")
                        
                except Exception as e:
                    st.error(f"An error occurred: {e}")

# ==========================================
# TOOL 2: TALLY LEDGER TO SUMMARY
# ==========================================
with tab2:
    st.header("Tally Ledger to Summary")
    uploaded_tally = st.file_uploader("Upload Tally Export (Excel)", type=["xlsx", "xls"], key="t2")
    
    if uploaded_tally:
        if st.button("Generate Summary", key="btn2"):
            try:
                df = pd.read_excel(uploaded_tally, engine='openpyxl')
                
                parties_data = {}
                
                for idx, row in df.iterrows():
                    try:
                        # Helper to safely get value by index
                        def get_safe(row, idx):
                            return row.iloc[idx] if len(row) > idx else None

                        particulars = str(get_safe(row, 2)).strip() if pd.notna(get_safe(row, 2)) else ''
                        credit = get_safe(row, 5)
                        debit = get_safe(row, 4)
                        
                        if not particulars or particulars == 'nan' or particulars == '': continue
                        if any(skip in particulars for skip in ['Closing Balance', 'nan', '700224', 'Balance', 'Ledger Account', '1-Apr']): continue
                        
                        party = particulars.replace('(Rent)', '').replace('(Interest)', '').strip()
                        if not party or '(as per details)' in party or len(party) < 3 or party.isdigit(): continue
                        
                        amount = None
                        # Check Credit
                        if pd.notna(credit) and credit != '' and credit != 0:
                            try:
                                c_val = float(credit)
                                if c_val > 0: amount = c_val
                            except: pass
                        
                        # Check Debit if no Credit
                        if amount is None and pd.notna(debit) and debit != '' and debit != 0:
                            try:
                                d_val = float(debit)
                                if d_val > 0: amount = d_val
                            except: pass
                            
                        if amount and amount > 0:
                            if party in parties_data:
                                parties_data[party] += amount
                            else:
                                parties_data[party] = amount
                    except: pass
                
                # Create Output Excel
                output = io.BytesIO()
                wb = Workbook()
                ws = wb.active
                ws.title = "Party Summary"
                
                # Headers
                ws['A1'] = "Name of Party"
                ws['B1'] = "Amount as per Books"
                
                # Styles
                header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
                header_font = Font(bold=True, color="FFFFFF")
                border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                
                for cell in ['A1', 'B1']:
                    ws[cell].fill = header_fill
                    ws[cell].font = header_font
                    ws[cell].border = border
                
                row_num = 2
                total_amount = 0
                
                for party in sorted(parties_data.keys()):
                    amt = parties_data[party]
                    total_amount += amt
                    ws[f'A{row_num}'] = party
                    ws[f'B{row_num}'] = amt
                    ws[f'B{row_num}'].number_format = '#,##0.00'
                    ws[f'A{row_num}'].border = border
                    ws[f'B{row_num}'].border = border
                    row_num += 1
                    
                # Total Row
                ws[f'A{row_num}'] = "TOTAL"
                ws[f'B{row_num}'] = total_amount
                ws[f'A{row_num}'].font = Font(bold=True)
                ws[f'B{row_num}'].font = Font(bold=True)
                ws[f'B{row_num}'].number_format = '#,##0.00'
                
                ws.column_dimensions['A'].width = 50
                ws.column_dimensions['B'].width = 20
                
                wb.save(output)
                output.seek(0)
                
                st.success(f"Processed! Found {len(parties_data)} unique parties.")
                st.download_button(label="Download Party Summary", data=output, file_name="Party_Summary.xlsx")
                
            except Exception as e:
                st.error(f"Error processing file: {e}")

# ==========================================
# TOOL 3: RECONCILIATION
# ==========================================
with tab3:
    st.header("Reconcile Books vs 26AS")
    col1, col2 = st.columns(2)
    with col1:
        file_books = st.file_uploader("Upload Books Summary (Excel)", type="xlsx", key="f1")
    with col2:
        file_26as = st.file_uploader("Upload 26AS Summary (Excel)", type="xlsx", key="f2")
        
    threshold = st.slider("Fuzzy Match Sensitivity", 50, 100, 75, help="Lower number matches loosely, Higher number matches strictly.")

    if file_books and file_26as:
        if st.button("Run Reconciliation", key="btn3"):
            try:
                df_books = pd.read_excel(file_books)
                df_26as = pd.read_excel(file_26as)
                
                # Standardize Columns
                df_books.columns = ['Name of Party', 'Amount in Books']
                df_26as.columns = ['Name of Party', 'Amount in 26AS']
                
                # Clean Data
                df_books = df_books[df_books['Name of Party'].str.upper() != 'TOTAL'].dropna()
                df_26as = df_26as[df_26as['Name of Party'].str.upper() != 'TOTAL'].dropna()
                
                # Fuzzy Logic
                matched_pairs = {}
                matched_26as_indices = set()
                
                for idx_b, party_b in enumerate(df_books['Name of Party']):
                    best_match = None
                    best_score = 0
                    
                    for idx_a, party_a in enumerate(df_26as['Name of Party']):
                        score = fuzz.token_set_ratio(str(party_b).upper(), str(party_a).upper())
                        if score > best_score:
                            best_score = score
                            best_match = idx_a
                    
                    if best_score >= threshold:
                        matched_pairs[idx_b] = best_match
                        matched_26as_indices.add(best_match)

                # Build Result
                reco_data = []
                
                # 1. Matches
                for idx_b, idx_a in matched_pairs.items():
                    # Use .iloc explicitly to avoid index confusion
                    b_row = df_books.iloc[idx_b]
                    a_row = df_26as.iloc[idx_a]
                    reco_data.append({
                        'Name of Party': b_row['Name of Party'],
                        'Amount in Books': b_row['Amount in Books'],
                        'Amount in 26AS': a_row['Amount in 26AS'],
                        'Difference': b_row['Amount in Books'] - a_row['Amount in 26AS']
                    })
                
                # 2. Only in Books
                for idx_b in range(len(df_books)):
                    if idx_b not in matched_pairs:
                        row = df_books.iloc[idx_b]
                        reco_data.append({
                            'Name of Party': row['Name of Party'],
                            'Amount in Books': row['Amount in Books'],
                            'Amount in 26AS': 0,
                            'Difference': row['Amount in Books']
                        })
                        
                # 3. Only in 26AS
                for idx_a in range(len(df_26as)):
                    if idx_a not in matched_26as_indices:
                        row = df_26as.iloc[idx_a]
                        reco_data.append({
                            'Name of Party': row['Name of Party'],
                            'Amount in Books': 0,
                            'Amount in 26AS': row['Amount in 26AS'],
                            'Difference': -row['Amount in 26AS']
                        })
                        
                final_df = pd.DataFrame(reco_data).sort_values('Name of Party')
                
                # Add Totals
                total_row = pd.DataFrame({
                    'Name of Party': ['TOTAL'],
                    'Amount in Books': [final_df['Amount in Books'].sum()],
                    'Amount in 26AS': [final_df['Amount in 26AS'].sum()],
                    'Difference': [final_df['Difference'].sum()]
                })
                final_df = pd.concat([final_df, total_row], ignore_index=True)
                
                # Excel Formatting
                output = io.BytesIO()
                final_df.to_excel(output, index=False, sheet_name='Reconciliation')
                
                # Apply Colors (Using openpyxl on the memory buffer)
                output.seek(0) # Rewind
                wb = load_workbook(output)
                ws = wb.active
                
                # Styles
                matched_fill = PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid') # Green
                diff_fill = PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid') # Red
                total_fill = PatternFill(start_color='FFEB9C', end_color='FFEB9C', fill_type='solid') # Yellow
                
                for row in ws.iter_rows(min_row=2):
                    party = row[0].value
                    diff = row[3].value
                    
                    fill_to_apply = None
                    if party == 'TOTAL':
                        fill_to_apply = total_fill
                    elif diff == 0 or abs(diff) < 1.0: # Floating point tolerance
                        fill_to_apply = matched_fill
                    else:
                        fill_to_apply = diff_fill
                        
                    if fill_to_apply:
                        for cell in row:
                            cell.fill = fill_to_apply
                            
                # Save final
                final_output = io.BytesIO()
                wb.save(final_output)
                final_output.seek(0)
                
                st.success("Reconciliation Complete!")
                st.download_button("Download Reconciliation Statement", data=final_output, file_name="Reconciliation_Statement.xlsx")
                
            except Exception as e:
                st.error(f"Reconciliation Error: {e}")