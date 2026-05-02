import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import os
import openpyxl
import sqlite3
from copy import copy
from datetime import datetime, date

from openpyxl.cell.text import InlineFont
from openpyxl.cell.rich_text import TextBlock, CellRichText
from openpyxl.styles.colors import Color
from openpyxl.styles import Font, Border, Side

# --- THE DEFAULT TEMPLATE PATH ---
TEMPLATE_PATH = "POTemplate.xlsx"

# --- DEFINING THE BORDER FOR AUTO-FILLED DATA ---
thin_border = Border(
    left=Side(style='thin', color='A5A5A5'), 
    right=Side(style='thin', color='A5A5A5'), 
    top=Side(style='thin', color='A5A5A5'), 
    bottom=Side(style='thin', color='A5A5A5')
)

# --- FACTORY NAME FORMATTER ---
def format_factory_name(name):
    if not name: return name
    name = str(name)
    name = re.sub(r'([a-zA-Z])\(', r'\1 (', name)
    name = re.sub(r'\)([a-zA-Z])', r') \1', name)
    name = re.sub(r'(Trax)(Apparel)', r'\1 \2', name, flags=re.IGNORECASE)
    return name.strip()

# --- RICH TEXT FORMATTER FOR COLUMN AK ---
def format_ak_column(text_val):
    if not text_val: return text_val
    text_str = str(text_val)
    blue_color = Color(rgb="FF0070C0")
    red_color = Color(rgb="FFFF0000")
    blue_font = InlineFont(rFont="Calibri", sz=11, color=blue_color)
    red_font = InlineFont(rFont="Calibri", sz=11, color=red_color)

    if "between 50% to 74.9%" in text_str or "between 0% to 49.9%" in text_str:
        parts = re.split(r'(Goose down|Duck down)', text_str)
        rich_text_elements = []
        for part in parts:
            if not part: continue
            if part in ["Goose down", "Duck down"]:
                rich_text_elements.append(TextBlock(red_font, part))
            else:
                rich_text_elements.append(TextBlock(blue_font, part))
        return CellRichText(*rich_text_elements)

    if "Filling:" in text_str:
        return CellRichText(TextBlock(blue_font, text_str))

    return text_str

# --- 1. THE EXTRACTION LOGIC ---
def parse_purchase_order(file_object):
    po_details = {
        "PO Number": "Not Found", "Date": "Not Found", "Header Ref Code": "Not Found",
        "Item ID": "Not Found", "Remark Size": "Not Found", "Order": "OTHERS", 
        "Style Name": "Not Found", "Age Group": "Not Found", "Gender": "Not Found", 
        "Garment Type": "Not Found", "Product Group": "Apparel", 
        "Factory Line": "Trax Apparel (Cambodia) Co.,Ltd.", "Product Division": "APP", 
        "Leather Logo": "No", "COO": "MADE IN CAMBODIA", "Season": "Not Found"
    }
    items_data = []

    try:
        with pdfplumber.open(file_object) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text(x_tolerance=1.2, y_tolerance=3)
                if text:
                    full_text += text + "\n"

            po_match = re.search(r"K[A-Z]{2,5}\d{6,}", full_text)
            if po_match: 
                po_details["PO Number"] = po_match.group(0)
            else:
                po_details["PO Number"] = file_object.name.replace(".pdf", "").replace(".PDF", "")

            item_match = re.search(r"(\d{8}-[A-Z]+-[A-Z0-9]+)", full_text)
            if item_match:
                clean_code = item_match.group(1).split("-")[-1] 
                po_details["Item ID"] = clean_code
                po_details["Header Ref Code"] = clean_code

            size_match = re.search(r"SIZE\s*PAGE\s*([^\n]+)", full_text, re.IGNORECASE)
            if size_match: 
                po_details["Remark Size"] = size_match.group(1).strip()

            # --- THE FINAL, BULLETPROOF MARKET EXTRACTION STRATEGY ---
            # 1. Find EVERY time "Order" is followed by a word (ignoring punctuation/spaces)
            market_matches = re.findall(r"Order[\s:;-]*([A-Za-z]+)", full_text, re.IGNORECASE)
            
            valid_market = "OTHERS"
            if market_matches:
                # 2. Filter out known bad grabs like "Order Quantity", "Order Form", or "Purchase Order" (where "Order" is the match)
                for match in reversed(market_matches): # Search backwards, the real one is usually at the end
                    clean_match = match.strip().upper()
                    if clean_match not in ["QUANTITY", "FORM", "DATE", "NUMBER", "NO", "PCS"]:
                        valid_market = clean_match
                        break
            
            po_details["Order"] = valid_market

            garment_patterns = {
                # "Order" is handled above now!
                "Style Name": r"Style\s*Name[\s:;]*(.*?)(?=;|Age|Gender|Garment|Product|COO)",
                "Age Group": r"Age\s*Group[\s:;]*(.*?)(?=;|Gender|Garment|Product|COO)",
                "Gender": r"Gender[\s:;]*(.*?)(?=;|Garment|Product|COO)",
                "Garment Type": r"Garment\s*Type[\s:;]*(.*?)(?=;|Product|Factory|COO)",
                "Product Group": r"Product\s*Group[/\w\s]*[\s:;]*(.*?)(?=;|Factory|product)",
                "Factory Line": r"Factory\s*Line\s*name[\s:;]*(.*?)(?=;|product|Made|COO)",
                "Product Division": r"product\s*division[\s:;]*(.*?)(?=;|Made|COO)",
                "Leather Logo": r"Made\s*by\s*leather[\s:;]*(.*?)(?=;|COO)",
                "COO": r"COO[\s:;]*(.*?)(?=;|\n|001A)",
                "Season": r"Season[\s:;]*([A-Z0-9]+)"
            }

            for key, pattern in garment_patterns.items():
                match = re.search(pattern, full_text, re.IGNORECASE | re.DOTALL)
                if match:
                    # Clean up the extracted text safely
                    clean_val = match.group(1).replace('\n', ' ').strip()
                    # Condense multiple accidental spaces down to a single space
                    clean_val = re.sub(r'\s+', ' ', clean_val)
                    clean_val = clean_val.rstrip(';/').strip()
                    
                    if key == "Order":
                        # Scrub out any accidental grab of "Order" or "Market"
                        clean_val = clean_val.replace("Order", "").replace("Market", "").strip()
                    po_details[key] = clean_val

            lines = full_text.split('\n')
            for line in lines:
                if "PCS" in line:
                    p = line.split()
                    if "PCS" in p:
                        if len(p) > 2 and (p[1] == "NO" or p[2] == "NO"): continue
                        pcs_idx = p.index("PCS")
                        try:
                            items_data.append({
                                "PO Number": po_details["PO Number"], 
                                "Material": p[0], "Color": p[1], "Size": p[2], "Sub-Code": p[3],
                                "Qty": float(p[pcs_idx - 1]), "Unit": "PCS",
                                "Price": float(p[pcs_idx + 1]), "Amount": float(p[-2]), "Delivery": p[-1]
                            })
                        except (ValueError, IndexError): continue
                            
    except Exception as e:
        st.error(f"Error reading PDF: {e}")
        return None, None
    return po_details, pd.DataFrame(items_data)

# --- 2. THE EXACT-STYLE TEMPLATE POPULATOR ---
def populate_existing_template(template_path, df_merged):
    if df_merged.empty: return None
    workbook = openpyxl.load_workbook(template_path)
    sheet = workbook.active
    
    header_row = 19 
    for r in range(1, 50): 
        if "*PO# number" in str(sheet.cell(row=r, column=1).value):
            header_row = r
            break
    start_row = header_row + 1 

    max_template_col = 37 
    for c in range(100, 0, -1):
        if sheet.cell(row=header_row, column=c).value is not None and str(sheet.cell(row=header_row, column=c).value).strip() != "":
            max_template_col = c
            break

    for r in range(1, start_row + 1):
        ak_cell = sheet.cell(row=r, column=37)
        if ak_cell.value:
            ak_cell.value = format_ak_column(ak_cell.value)
            ak_cell.font = Font(name="Calibri", size=11)

    for index, row in df_merged.iterrows():
        current_row = start_row + index
        
        mapping = {
            1: row["*PO# number"], 2: row["*Style#"], 3: row["*Article#"], 4: row["*Product group / type "],
            5: row["*Style name"], 6: row["*COO"], 7: row["*Brand"], 
            8: row["*Factory line name "], 10: row["Leather logo (optional)"], 11: row["*Age group"], 
            12: row["*Size page"], 13: row["*Garment type"], 14: row["*Gender"], 15: row["*Product division "],
            19: row["*Sourcing size"], 22: row["*Order Quantity"], 23: row["*Order Market "], 
            24: row["*Order Type"], 25: row["*Season"]
        }

        for col_num in range(1, max_template_col + 1):
            target_cell = sheet.cell(row=current_row, column=col_num)
            
            if current_row != start_row: 
                template_cell = sheet.cell(row=start_row, column=col_num)
                if template_cell.has_style:
                    target_cell.fill = copy(template_cell.fill)
                    target_cell.number_format = copy(template_cell.number_format)
                    target_cell.protection = copy(template_cell.protection)
                    target_cell.alignment = copy(template_cell.alignment)
                    target_cell.font = copy(template_cell.font)
                if col_num not in mapping:
                    target_cell.value = template_cell.value

            if col_num in mapping:
                target_cell.value = mapping[col_num]
            
            target_cell.border = thin_border
            if col_num == 22: 
                target_cell.number_format = '0.00'
            if col_num == 37 and target_cell.value:
                target_cell.value = format_ak_column(target_cell.value)
                target_cell.font = Font(name="Calibri", size=11)

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output

# --- 3. THE UI & NAVIGATION ---
st.set_page_config(page_title="PO Data Extractor", page_icon=":material/description:", layout="wide")

if "preview_df" not in st.session_state:
    st.session_state.preview_df = None
if "report_name_default" not in st.session_state:
    st.session_state.report_name_default = "Extracted_PO_Data"

st.sidebar.title(":material/dashboard: Control Panel")
menu = st.sidebar.radio("Navigate Systems:", [
    ":material/download: 1. Data Extraction", 
    ":material/settings: 2. Data Processing & Export", 
    ":material/bar_chart: 3. Reports"
])

if menu == ":material/download: 1. Data Extraction":
    st.title(":material/download: Data Extraction")
    st.markdown("Upload your PO PDFs. The system will parse the documents and structure the data automatically.")

    uploaded_pdfs = st.file_uploader("Upload PO PDFs for Extraction", type=["pdf"], accept_multiple_files=True)

    if uploaded_pdfs:
        all_po_details, all_items_data = [], []
        with st.spinner('Analyzing Purchase Orders...'):
            for file in uploaded_pdfs:
                details, df_items = parse_purchase_order(file)
                if details:
                    all_po_details.append(details)
                    if not df_items.empty: all_items_data.append(df_items)
        
        if all_po_details:
            master_details_df = pd.DataFrame(all_po_details)
            master_items_df = pd.concat(all_items_data, ignore_index=True) if all_items_data else pd.DataFrame()
            
            preview_rows = []
            df_merged_raw = pd.merge(master_items_df, master_details_df, on="PO Number", how="left")
            
            for _, row in df_merged_raw.iterrows():
                age_group_val = str(row["Age Group"]).capitalize() if pd.notna(row["Age Group"]) and str(row["Age Group"]).upper() != "NOT FOUND" else ""
                
                # Dynamic Market Check Cleanup
                raw_market = str(row["Order"]).strip()
                if not raw_market or raw_market.upper() in ["NOT FOUND", "OTHERS", "", "ORDER"]:
                    market_found = "OTHERS"
                else:
                    # Clean up the capitalization to be uppercase
                    market_found = raw_market.upper()
                
                preview_rows.append({
                    "*PO# number": row["PO Number"],
                    "*Style#": row["Item ID"],
                    "*Article#": row["Sub-Code"],
                    "*Product group / type ": row["Product Group"],
                    "*Style name": row["Style Name"],
                    "*COO": "MADE IN CAMBODIA",
                    "*Brand": "Adidas",
                    "*Factory line name ": format_factory_name(row["Factory Line"]),
                    "Leather logo (optional)": row["Leather Logo"],
                    "*Age group": age_group_val,
                    "*Size page": row["Remark Size"],
                    "*Garment type": row["Garment Type"],
                    "*Gender": row["Gender"],
                    "*Product division ": row["Product Division"],
                    "*Sourcing size": row["Size"],
                    "*Order Quantity": row["Qty"],
                    "*Order Market ": market_found,
                    "*Order Type": "BULK",
                    "*Season": row["Season"]
                })
            
            st.session_state.preview_df = pd.DataFrame(preview_rows)
            st.success(":material/check_circle: Analysis Complete! Preview the structured data in the 'Data Processing & Export' tab or save it to the Database below.")

            st.markdown("---")
            st.subheader(":material/save: Save Data to Database")
            if len(uploaded_pdfs) == 1:
                first_po = master_details_df.iloc[0]['PO Number']
                default_report_name = f"{first_po}_Report"
            else:
                default_report_name = f"Batch_PO_Report_{len(uploaded_pdfs)}_Files"
            
            st.session_state.report_name_default = default_report_name
            report_name = st.text_input("Data Batch ID:", value=st.session_state.report_name_default)
            
            if st.button(":material/database: Initialize Data Injection"):
                db_df = st.session_state.preview_df.copy()
                db_df.insert(0, "Uploaded_Date", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                db_df.insert(0, "Report_Name", report_name)
                
                try:
                    conn = sqlite3.connect('po_database.db')
                    db_df.to_sql('uploaded_reports', conn, if_exists='append', index=False)
                    conn.close()
                    st.success(f":material/check_circle: Successfully saved {len(db_df)} records into the Database!")
                except Exception as e:
                    st.error(f":material/error: Database Error: {e}")

elif menu == ":material/settings: 2. Data Processing & Export":
    st.title(":material/settings: Data Processing & Template Generation")
    
    if st.session_state.preview_df is not None and not st.session_state.preview_df.empty:
        
        # --- NEW FIX: SORTING FEATURE ---
        st.markdown("### :material/sort: Sort Your Data Before Export")
        sort_c1, sort_c2 = st.columns(2)
        sort_column = sort_c1.selectbox("Select Column to Sort By:", ["(No Sorting)"] + list(st.session_state.preview_df.columns))
        sort_order = sort_c2.radio("Sort Order:", ["Ascending :material/arrow_upward:", "Descending :material/arrow_downward:"], horizontal=True)
        
        # Apply sorting if a column is selected and reset index for Excel generation
        export_df = st.session_state.preview_df.copy()
        if sort_column != "(No Sorting)":
            is_ascending = (sort_order == "Ascending :material/arrow_upward:")
            # reset_index(drop=True) ensures no blank rows are skipped in the Excel template
            export_df = export_df.sort_values(by=sort_column, ascending=is_ascending).reset_index(drop=True)

        st.markdown("Review the extracted data before generating the final output:")
        st.dataframe(export_df, use_container_width=True)
        
        st.markdown("---")
        if not os.path.exists(TEMPLATE_PATH):
            st.error(f":material/error: Error: Could not find `{TEMPLATE_PATH}`. Please ensure your template file is uploaded.")
        else:
            with st.spinner("Compiling formatted Excel file..."):
                final_excel_file = populate_existing_template(TEMPLATE_PATH, export_df)
                
            if final_excel_file:
                st.download_button(
                    label=":material/file_download: Download Compiled Excel Matrix", 
                    data=final_excel_file, 
                    file_name=f"{st.session_state.report_name_default}.xlsx", 
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
    else:
        st.info(":material/info: Awaiting input data. Please navigate to **:material/download: 1. Data Extraction** to initialize analysis.")

elif menu == ":material/bar_chart: 3. Reports":
    st.title(":material/bar_chart: Reports & Dashboard")
    st.markdown("View extraction statistics and query the local database.")
    
    if os.path.exists('po_database.db'):
        try:
            conn = sqlite3.connect('po_database.db')
            history_df = pd.read_sql("SELECT * FROM uploaded_reports ORDER BY Uploaded_Date DESC", conn)
            conn.close()

            # Convert string dates to actual datetime objects for analysis
            history_df['Uploaded_Date_DT'] = pd.to_datetime(history_df['Uploaded_Date'])

            # --- DASHBOARD METRICS ---
            today = datetime.now().date()
            start_of_month = today.replace(day=1)

            # Isolate DataFrames by time periods
            df_today = history_df[history_df['Uploaded_Date_DT'].dt.date == today]
            df_month = history_df[(history_df['Uploaded_Date_DT'].dt.date >= start_of_month) & (history_df['Uploaded_Date_DT'].dt.date <= today)]

            # Calculate Unique POs (Files)
            today_po_count = df_today['*PO# number'].nunique()
            month_po_count = df_month['*PO# number'].nunique()
            total_po_count = history_df['*PO# number'].nunique()

            # Calculate Total Records (Line Items)
            today_count = len(df_today)
            month_count = len(df_month)
            total_count = len(history_df)

            st.markdown("### :material/monitoring: Extraction Dashboard")
            
            st.markdown("**::material/folder: Unique PO Files Processed**")
            m1, m2, m3 = st.columns(3)
            m1.metric("POs Today", today_po_count)
            m2.metric("POs This Month", month_po_count)
            m3.metric("Total POs in Database", total_po_count)
            
            st.markdown("**:material/list_alt: Line Items Extracted**")
            m4, m5, m6 = st.columns(3)
            m4.metric("Records Today", today_count)
            m5.metric("Records This Month", month_count)
            m6.metric("Total Records in Database", total_count)

            st.markdown("---")
            
            # --- FILTERS ---
            st.markdown("### :material/tune: Query Parameters")
            col1, col2, col3, col4 = st.columns(4)
            
            po_filter = col1.text_input("Filter by PO Number")
            style_filter = col2.text_input("Filter by Style Name")
            
            market_options = history_df["*Order Market "].dropna().unique().tolist()
            market_filter = col3.multiselect("Filter by Market", options=market_options)

            # Date Range Filter (Default to Current Month)
            date_range = col4.date_input("Filter by Date Range", value=(start_of_month, today))
            
            filtered_df = history_df.copy()

            if po_filter:
                filtered_df = filtered_df[filtered_df["*PO# number"].str.contains(po_filter, case=False, na=False)]
            if style_filter:
                filtered_df = filtered_df[filtered_df["*Style name"].str.contains(style_filter, case=False, na=False)]
            if market_filter:
                filtered_df = filtered_df[filtered_df["*Order Market "].isin(market_filter)]

            # Apply Date Filter based on user selection
            if isinstance(date_range, tuple) and len(date_range) == 2:
                start_d, end_d = date_range
                filtered_df = filtered_df[(filtered_df['Uploaded_Date_DT'].dt.date >= start_d) & (filtered_df['Uploaded_Date_DT'].dt.date <= end_d)]
            elif isinstance(date_range, tuple) and len(date_range) == 1:
                start_d = date_range[0]
                filtered_df = filtered_df[filtered_df['Uploaded_Date_DT'].dt.date == start_d]
            elif isinstance(date_range, date): 
                filtered_df = filtered_df[filtered_df['Uploaded_Date_DT'].dt.date == date_range]

            # Clean up the temporary datetime column before displaying
            if 'Uploaded_Date_DT' in filtered_df.columns:
                filtered_df = filtered_df.drop(columns=['Uploaded_Date_DT'])
                
            st.markdown(f"**Showing {len(filtered_df)} records (from {filtered_df['*PO# number'].nunique()} unique POs)**")
            st.dataframe(filtered_df, use_container_width=True)
            
        except Exception as e:
            st.error(f":material/error: Error reading database: {e}")
    else:
        st.info(":material/info: The database is empty. Upload and process some POs to begin storing data.")