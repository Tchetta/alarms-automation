import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
import io
import zipfile

st.set_page_config(page_title="Power Alarms Processor", page_icon="⚡", layout="centered")

st.title("⚡ Power Alarms Automation Dashboard")
st.write("Upload your fresh dashboard export. The application will automatically synchronize with your master reference database.")

# ==========================================
# 1. GITHUB MASTER REFERENCE CONFIGURATION
# ==========================================
GITHUB_RAW_REF_URL = "https://raw.githubusercontent.com/Tchetta/alarms-automation/main/ref1.xlsx"

# 2. FILE UPLOAD INTERFACE
uploaded_alarm = st.file_uploader("Upload Power Alarm Export (.xlsx, .csv, or .zip)", type=["xlsx", "csv", "zip"])

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

if uploaded_alarm:
    if st.button("🚀 Generate Formatted Alarm Report"):
        with st.spinner("Fetching database reference from cloud repository and cleaning data..."):
            
            # ==========================================
            # 3. FETCH & VALIDATE GITHUB MASTER REFERENCE
            # ==========================================
            try:
                st.info("🔄 Connecting to cloud master database reference...")
                df_ref = pd.read_excel(GITHUB_RAW_REF_URL)
                
                # Strip spaces from column headers
                df_ref.columns = [str(c).strip() for c in df_ref.columns]
                
                # Verify structural components
                required_ref_cols = ['Site ID', 'Site Name', 'Power Co', 'Children', 'Backbone Site']
                missing_ref_cols = [col for col in required_ref_cols if col not in df_ref.columns]
                
                if missing_ref_cols:
                    st.error(f"❌ Master reference file structure mismatch! Missing columns: {missing_ref_cols}")
                    st.stop()
                
                # Data cleaning for Master Reference
                df_ref['Site ID'] = pd.to_numeric(df_ref['Site ID'].astype(str).str.strip(), errors='coerce')
                df_ref_clean = df_ref.dropna(subset=['Site ID']).copy()
                df_ref_clean['Site ID'] = df_ref_clean['Site ID'].astype(int)
                
                # Create dictionaries for fast row lookups
                power_co_map = df_ref_clean.set_index('Site ID')['Power Co'].to_dict()
                children_map = df_ref_clean.set_index('Site ID')['Children'].to_dict()
                backbone_map = df_ref_clean.set_index('Site ID')['Backbone Site'].to_dict()
                
                st.success("✅ Master reference synchronized successfully.")
                
            except Exception as ref_err:
                st.error(f"❌ Failed to download or parse master reference from GitHub: {ref_err}")
                st.stop()

            # ==========================================
            # 4. PARSE LOCAL ALARM EXPORT
            # ==========================================
            try:
                alarm_file_stream = uploaded_alarm
                file_name = uploaded_alarm.name.lower()

                if file_name.endswith('.zip'):
                    with zipfile.ZipFile(uploaded_alarm) as z:
                        zip_files = z.namelist()
                        if not zip_files:
                            st.error("❌ The uploaded .zip archive contains no files.")
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
                        if "not a zip file" in str(excel_err).lower() or "bad zip file" in str(excel_err).lower():
                            html_tables = pd.read_html(alarm_file_stream)
                            if html_tables:
                                df_alarm = html_tables[0]
                            else:
                                raise excel_err
                        else:
                            raise excel_err
                            
                df_alarm.columns = [str(c).strip() for c in df_alarm.columns]
                
            except Exception as read_err:
                st.error(f"❌ Could not interpret the uploaded Alarm Export: {read_err}")
                st.stop()

            # ==========================================
            # 5. DATA TRANSFORMATION & CORRELATION PIPELINE
            # ==========================================
            try:
                required_alarm_cols = ['Site ID', 'Site Name', 'Ticket ID', 'Alarm Name', 'First Occurred On', 'Duration(hh:mm:ss)', 'Last Occurred On']
                missing_alarm_cols = [col for col in required_alarm_cols if col not in df_alarm.columns]
                
                if missing_alarm_cols:
                    st.error(f"❌ Alarm upload missing expected raw columns: {missing_alarm_cols}")
                    st.stop()
                
                df_alarm['Site ID'] = pd.to_numeric(df_alarm['Site ID'].astype(str).str.strip(), errors='coerce')
                df_clean = df_alarm.dropna(subset=['Site ID']).copy()
                df_clean['Site ID'] = df_clean['Site ID'].astype(int)

                df_report = df_clean[required_alarm_cols].copy()
                df_report['Power Owner'] = df_report['Site ID'].map(power_co_map).fillna("IHS")
                
                # Dynamic Logic Condition: Backbone prioritization rules
                def determine_prediction(site_id):
                    is_backbone = str(backbone_map.get(site_id, '')).strip().lower() == 'yes'
                    kids_val = children_map.get(site_id)
                    
                    try:
                        kids_count = int(float(kids_val)) if pd.notna(kids_val) and str(kids_val).strip() != "" else 0
                    except ValueError:
                        kids_count = 0
                    
                    if is_backbone and kids_count < 5:
                        return "BACKBONE site"
                    elif kids_count > 0:
                        return f"{kids_count} sites will be affected"
                    return ""
                
                df_report['Prediction'] = df_report['Site ID'].apply(determine_prediction)

                final_column_order = [
                    'Site ID', 'Site Name', 'Power Owner', 'Prediction', 
                    'Ticket ID', 'Alarm Name', 'First Occurred On', 
                    'Duration(hh:mm:ss)', 'Last Occurred On'
                ]
                df_final = df_report[final_column_order].sort_values(by='Site ID')

            except Exception as processing_err:
                st.error(f"❌ Data alignment failed: {processing_err}")
                st.stop()

            # ==========================================
            # 6. ENGINE PRESENTATION & STYLING (openpyxl)
            # ==========================================
            try:
                wb = Workbook()
                ws = wb.active
                ws.title = "Power Alarms Status"
                ws.views.sheetView[0].showGridLines = True

                # --- STEP A: GENERATE MERGED TITLE BANNER ---
                ws.merge_cells("A1:I1")
                title_cell = ws["A1"]
                title_cell.value = "Power Alarms on TIS"
                
                # Set unified orange fill matching your target theme
                orange_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
                
                title_cell.fill = orange_fill
                title_cell.font = Font(name="Calibri", size=14, bold=True, color="000000")
                title_cell.alignment = Alignment(horizontal="center", vertical="center")
                ws.row_dimensions[1].height = 30

                # --- STEP B: APPEND DATA UNDER THE TITLE ---
                ws.append(final_column_order) # This lands on Row 2
                for r in dataframe_to_rows(df_final, index=False, header=False):
                    ws.append(r)

                # Style configurations
                header_font = Font(name="Calibri", size=11, bold=True, color="000000")
                data_font = Font(name="Calibri", size=11, color="000000")
                
                black_side = Side(border_style="thin", color="000000")
                grid_border = Border(left=black_side, right=black_side, top=black_side, bottom=black_side)

                # Set auto filter parameters across Row 2 boundaries
                ws.auto_filter.ref = f"A2:I{ws.max_row}"
                ws.row_dimensions[2].height = 24

                # Apply crisp full black borders to the merged title row block components
                for col in range(1, 10):
                    ws.cell(row=1, column=col).border = grid_border

                # Style table headers (Row 2) - Using matching orange fill color layout
                for cell in ws[2]:
                    cell.fill = orange_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = grid_border

                # Style data matrix starting down at Row 3
                for row in range(3, ws.max_row + 1):
                    ws.row_dimensions[row].height = 18
                    for col in range(1, 10):
                        cell = ws.cell(row=row, column=col)
                        cell.font = data_font
                        cell.border = grid_border
                        
                        if col in [1, 5, 7, 8, 9]:
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        else:
                            cell.alignment = Alignment(horizontal="left", vertical="center")

                # Track column widths safely across merged rows using explicit index tracking
                for col_idx in range(1, 10):
                    max_len = 0
                    for row in range(2, ws.max_row + 1):  # Read length starting from Header row downwards
                        val = str(ws.cell(row=row, column=col_idx).value or '')
                        if len(val) > max_len:
                            max_len = len(val)
                    
                    col_letter = get_column_letter(col_idx)
                    ws.column_dimensions[col_letter].width = max(max_len + 4, 14)

                # Process production timestamps
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
                st.success("Report successfully generated!")
                
                st.download_button(
                    label=f"📥 Download {filename}",
                    data=output_buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as excel_layout_err:
                st.error(f"❌ Spreadsheet rendering or styling crash: {excel_layout_err}")
                
