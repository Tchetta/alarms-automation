import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime, timedelta
import io
import zipfile

st.set_page_config(page_title="Power Alarms Processor", page_icon="⚡", layout="centered")

st.title("⚡ Power Alarms Automation Dashboard")
st.write("Upload your fresh dashboard export and reference files to build your formatted report.")

# 1. FILE UPLOAD INTERFACE
uploaded_alarm = st.file_uploader("1. Upload Power Alarm Export (.xlsx, .csv, or .zip)", type=["xlsx", "csv", "zip"])
uploaded_pm = st.file_uploader("2. [OPTIONAL] Upload PM Plan Reference (Defaults to 'IHS' if missing)", type=["xlsx"])
uploaded_parent = st.file_uploader("3. Upload Parent Sites Reference (Parent-sites-Tx...)", type=["xlsx"])

def round_to_nearest_30_minutes(dt):
    """Rounds a datetime object up or down to the nearest 30-minute increment."""
    if pd.isna(dt):
        return datetime.now().replace(second=0, microsecond=0)
    minute = dt.minute
    remainder = minute % 30
    if remainder < 15:
        dt = dt - timedelta(minutes=remainder)
    else:
        dt = dt + timedelta(minutes=(30 - remainder))
    return dt.replace(second=0, microsecond=0)

# THE TRIGGER DEMANDS THE ALARM EXPORT AND PARENT SITES TO RUN
if uploaded_alarm and uploaded_parent:
    st.success("Required core files staged successfully.")
    
    if st.button("🚀 Generate Formatted Alarm Report"):
        with st.spinner("Processing network structures and applying styling configurations..."):
            
            # ==========================================
            # 2. INGEST DATA SOURCES
            # ==========================================
            
            # --- STEP A: PARSE ALARM EXPORT (WITH AUTOMATIC ZIP HANDLING) ---
            try:
                alarm_file_stream = uploaded_alarm
                file_name = uploaded_alarm.name.lower()

                if file_name.endswith('.zip'):
                    with zipfile.ZipFile(uploaded_alarm) as z:
                        zip_files = z.namelist()
                        if not zip_files:
                            st.error("❌ The uploaded .zip file is empty.")
                            st.stop()
                        
                        internal_filename = zip_files[0]
                        file_name = internal_filename.lower()
                        alarm_file_stream = io.BytesIO(z.read(internal_filename))

                if file_name.endswith('.csv'):
                    df_alarm = pd.read_csv(alarm_file_stream)
                else:
                    try:
                        df_alarm = pd.read_excel(alarm_file_stream)
                    except Exception as excel_err:
                        err_msg = str(excel_err).lower()
                        if "not a zip file" in err_msg or "bad zip file" in err_msg:
                            html_tables = pd.read_html(alarm_file_stream)
                            if html_tables:
                                df_alarm = html_tables[0]
                            else:
                                raise excel_err
                        else:
                            raise excel_err
            except Exception as read_err:
                st.error(f"❌ Could not parse the Alarm Export file: {read_err}.")
                st.stop()

            # --- STEP B: PARSE PARENT SITES REFERENCE ---
            try:
                df_parent = pd.read_excel(uploaded_parent, sheet_name=0)
            except Exception as parent_err:
                st.error(f"❌ Error reading the Parent Sites reference file: {parent_err}")
                st.stop()

            # --- STEP C: PROCESSING DATA PIPELINE ---
            try:
                # Core Alarm Cleanup: Strip formatting and cast to clean numeric integers
                df_alarm['Site ID'] = pd.to_numeric(df_alarm['Site ID'].astype(str).str.strip(), errors='coerce')
                df_alarm = df_alarm.dropna(subset=['Site ID'])
                df_alarm['Site ID'] = df_alarm['Site ID'].astype(int)

                # Parent Sites Cleanup: Standardize Column A tracking indices
                df_parent.iloc[:, 0] = pd.to_numeric(df_parent.iloc[:, 0].astype(str).str.strip(), errors='coerce')
                df_parent_clean = df_parent.dropna(subset=[df_parent.columns[0]]).copy()
                df_parent_clean.iloc[:, 0] = df_parent_clean.iloc[:, 0].astype(int)

                keep_cols = ['Site ID', 'Site Name', 'Ticket ID', 'Alarm Name', 'First Occurred On', 'Duration(hh:mm:ss)', 'Last Occurred On']
                df_clean = df_alarm[keep_cols].copy()

                # Build lookups from standard parent references
                parent_child_dict = df_parent_clean.set_index(df_parent_clean.columns[0])[df_parent_clean.columns[5]].to_dict()
                parent_owner_dict = df_parent_clean.set_index(df_parent_clean.columns[0])[df_parent_clean.columns[7]].to_dict()

                # ==========================================
                # HARDENED DYNAMIC POWER OWNER LOGIC
                # ==========================================
                if uploaded_pm:
                    try:
                        df_pm = pd.read_excel(uploaded_pm, sheet_name=0)
                        
                        # Fix data types: strip spaces, drop invalid rows, cast to uniform integers
                        df_pm.iloc[:, 0] = pd.to_numeric(df_pm.iloc[:, 0].astype(str).str.strip(), errors='coerce')
                        df_pm_clean = df_pm.dropna(subset=[df_pm.columns[0]]).copy()
                        df_pm_clean.iloc[:, 0] = df_pm_clean.iloc[:, 0].astype(int)
                        
                        # Build mapping dictionary out of pure, uniform types
                        pm_dict = df_pm_clean.set_index(df_pm_clean.columns[0])[df_pm_clean.columns[4]].to_dict()
                        
                        # Try mapping directly via PM Plan reference
                        df_clean['Power Owner'] = df_clean['Site ID'].map(pm_dict)
                        
                        # Fallback step 1: If PM Plan didn't have it, try mapping from Parent Sites Power Co
                        df_clean['Power Owner'] = df_clean['Power Owner'].fillna(df_clean['Site ID'].map(parent_owner_dict))
                        
                        # Fallback step 2: If neither file contains the site, default strictly to "IHS"
                        df_clean['Power Owner'] = df_clean['Power Owner'].fillna("IHS")
                        
                    except Exception as pm_err:
                        st.warning(f"⚠️ Could not cleanly match the PM Plan tracking columns ({pm_err}). Relying on Parent reference data details.")
                        df_clean['Power Owner'] = df_clean['Site ID'].map(parent_owner_dict).fillna("IHS")
                else:
                    st.info("ℹ️ PM Plan file missing. Mapping owners via Parent Sites 'Power Co' with 'IHS' fallback.")
                    df_clean['Power Owner'] = df_clean['Site ID'].map(parent_owner_dict).fillna("IHS")

                # Parse operational child site counts 
                temp_children = df_clean['Site ID'].map(parent_child_dict)
                df_clean['Prediction'] = temp_children.apply(
                    lambda x: f"{int(x)} sites will be affected" if pd.notna(x) and isinstance(x, (int, float)) else ""
                )

                final_column_order = [
                    'Site ID', 'Site Name', 'Power Owner', 'Prediction', 
                    'Ticket ID', 'Alarm Name', 'First Occurred On', 
                    'Duration(hh:mm:ss)', 'Last Occurred On'
                ]
                df_final = df_clean[final_column_order].copy()
                df_final = df_final.sort_values(by='Site ID')

            except KeyError as key_err:
                st.error(f"❌ Column Mapping Error: Expected headers are structural issues. Details: {key_err}")
                st.stop()
            except Exception as processing_err:
                st.error(f"❌ Error during data transformation: {processing_err}")
                st.stop()

            # --- STEP D: EXCEL PRESENTATION ENGINE (openpyxl) ---
            try:
                wb = Workbook()
                ws = wb.active
                ws.title = "Power Alarms Status"
                ws.views.sheetView[0].showGridLines = True

                ws.append(final_column_order)
                for r in dataframe_to_rows(df_final, index=False, header=False):
                    ws.append(r)

                header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
                header_font = Font(name="Calibri", size=11, bold=True, color="000000")
                data_font = Font(name="Calibri", size=11, color="000000")
                
                thin_side = Side(border_style="thin", color="D3D3D3")
                grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

                ws.auto_filter.ref = f"A1:I{ws.max_row}"

                for cell in ws[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = grid_border

                for row in range(2, ws.max_row + 1):
                    for col in range(1, 10):
                        cell = ws.cell(row=row, column=col)
                        cell.font = data_font
                        cell.border = grid_border
                        
                        if col in [1, 5, 7, 8, 9]:
                            cell.alignment = Alignment(horizontal="center")
                        else:
                            cell.alignment = Alignment(horizontal="left")

                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = col[0].column_letter
                    ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

                # --- TRACK FILENAME DATE DETAILS ---
                try:
                    dates_parsed = pd.to_datetime(df_final['Last Occurred On'], errors='coerce')
                    max_date = dates_parsed.max()
                    if pd.isna(max_date):
                        max_date = datetime.now()
                        
                    rounded_time = round_to_nearest_30_minutes(max_date)
                    time_string = rounded_time.strftime("%HH%M")
                except Exception:
                    time_string = datetime.now().strftime("%HH%M")

                filename = f"TIS_ALARMS{time_string}.xlsx"

                output_buffer = io.BytesIO()
                wb.save(output_buffer)
                output_buffer.seek(0)

                st.balloons()
                st.success(f"Report compiled successfully using reference timeline context: {time_string}!")
                
                st.download_button(
                    label=f"📥 Download {filename}",
                    data=output_buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as excel_layout_err:
                st.error(f"❌ Failed to build or style the output Excel workbook: {excel_layout_err}")
                
else:
    st.warning("Waiting for the Alarm Export and Parent Sites reference file to be uploaded...")
