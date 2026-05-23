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
st.write("Upload your fresh dashboard export. The application will automatically synchronize with your master reference database.")

# ==========================================
# 1. GITHUB MASTER REFERENCE CONFIGURATION
# ==========================================
# TODO: Replace this URL with your actual GitHub Raw link once uploaded.
# To get the raw link: Click your file on GitHub, click the "Raw" button, and copy that URL.
GITHUB_RAW_REF_URL = "https://raw.githubusercontent.com/Tchetta/alarms-automation/main/ref1.xlsx"

# 2. SINGLE FILE UPLOAD INTERFACE
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
                
                # Strip spaces from column headers to prevent accidental naming mismatches
                df_ref.columns = [str(c).strip() for c in df_ref.columns]
                
                # Strict structural verification matching your layout
                required_ref_cols = ['Site ID', 'Site Name', 'Power Co', 'Children']
                missing_ref_cols = [col for col in required_ref_cols if col not in df_ref.columns]
                
                if missing_ref_cols:
                    st.error(f"❌ Master reference file structure mismatch! Missing expected columns: {missing_ref_cols}")
                    st.stop()
                
                # Data cleaning for Master Reference
                df_ref['Site ID'] = pd.to_numeric(df_ref['Site ID'].astype(str).str.strip(), errors='coerce')
                df_ref_clean = df_ref.dropna(subset=['Site ID']).copy()
                df_ref_clean['Site ID'] = df_ref_clean['Site ID'].astype(int)
                
                # Map target columns down to lightweight key-value indices
                power_co_map = df_ref_clean.set_index('Site ID')['Power Co'].to_dict()
                children_map = df_ref_clean.set_index('Site ID')['Children'].to_dict()
                
                st.success("✅ Master reference synchronized successfully.")
                
            except Exception as ref_err:
                st.error(f"❌ Failed to download or parse master reference from GitHub: {ref_err}")
                st.info("💡 Double check that your GITHUB_RAW_REF_URL is correct and points to a public repository raw file stream.")
                st.stop()

            # ==========================================
            # 4. PARSE LOCAL ALARM EXPORT
            # ==========================================
            try:
                alarm_file_stream = uploaded_alarm
                file_name = uploaded_alarm.name.lower()

                # Unzip if file is compressed
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
                        # Fallback case for HTML tables disguised as Excel spreadsheets
                        if "not a zip file" in str(excel_err).lower() or "bad zip file" in str(excel_err).lower():
                            html_tables = pd.read_html(alarm_file_stream)
                            if html_tables:
                                df_alarm = html_tables[0]
                            else:
                                raise excel_err
                        else:
                            raise excel_err
                            
                # Strip spaces from raw alarm columns
                df_alarm.columns = [str(c).strip() for c in df_alarm.columns]
                
            except Exception as read_err:
                st.error(f"❌ Could not interpret the uploaded Alarm Export: {read_err}")
                st.stop()

            # ==========================================
            # 5. DATA TRANSFORMATION & CORRELATION PIPELINE
            # ==========================================
            try:
                # Target layout extraction structure
                required_alarm_cols = ['Site ID', 'Site Name', 'Ticket ID', 'Alarm Name', 'First Occurred On', 'Duration(hh:mm:ss)', 'Last Occurred On']
                missing_alarm_cols = [col for col in required_alarm_cols if col not in df_alarm.columns]
                
                if missing_alarm_cols:
                    st.error(f"❌ Alarm upload missing expected raw columns: {missing_alarm_cols}")
                    st.stop()
                
                # Standardize primary IDs
                df_alarm['Site ID'] = pd.to_numeric(df_alarm['Site ID'].astype(str).str.strip(), errors='coerce')
                df_clean = df_alarm.dropna(subset=['Site ID']).copy()
                df_clean['Site ID'] = df_clean['Site ID'].astype(int)

                # Pull primary attributes
                df_report = df_clean[required_alarm_cols].copy()

                # Perform cloud dictionary mappings
                df_report['Power Owner'] = df_report['Site ID'].map(power_co_map).fillna("IHS")
                
                # Process affected site calculations string safely
                def format_children(site_id):
                    val = children_map.get(site_id)
                    if pd.notna(val) and str(val).strip() != "" and float(val) > 0:
                        return f"{int(float(val))} sites will be affected"
                    return ""
                
                df_report['Prediction'] = df_report['Site ID'].apply(format_children)

                # Arrange output headers to match exact required production output order
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

                # Structural append
                ws.append(final_column_order)
                for r in dataframe_to_rows(df_final, index=False, header=False):
                    ws.append(r)

                # Typography & Palette configurations
                header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
                header_font = Font(name="Calibri", size=11, bold=True, color="000000")
                data_font = Font(name="Calibri", size=11, color="000000")
                
                # Hardened solid black production cell border grids
                black_side = Side(border_style="thin", color="000000")
                grid_border = Border(left=black_side, right=black_side, top=black_side, bottom=black_side)

                ws.auto_filter.ref = f"A1:I{ws.max_row}"

                # Style top row headers
                for cell in ws[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = grid_border

                # Style data rows
                for row in range(2, ws.max_row + 1):
                    for col in range(1, 10):
                        cell = ws.cell(row=row, column=col)
                        cell.font = data_font
                        cell.border = grid_border
                        
                        # Structured alignments
                        if col in [1, 5, 7, 8, 9]:
                            cell.alignment = Alignment(horizontal="center")
                        else:
                            cell.alignment = Alignment(horizontal="left")

                # Track context width padding dynamically
                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = col[0].column_letter
                    ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

                # Process file names by rounding timestamps
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

                # Save spreadsheet down to binary buffer stream
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
