import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime, timedelta
import io

st.set_page_config(page_title="Power Alarms Processor", page_icon="⚡", layout="centered")

st.title("⚡ Power Alarms Automation Dashboard")
st.write("Upload your fresh dashboard export and reference files to build your formatted report.")

# 1. FILE UPLOAD INTERFACE
uploaded_alarm = st.file_uploader("1. Upload Power Alarm Export (.xlsx or .csv)", type=["xlsx", "csv"])
# WE REMOVED THE 'required' STATUS FROM THE PM PLAN FILE LABELLING
uploaded_pm = st.file_uploader("2. [OPTIONAL] Upload PM Plan Reference (Defaults to 'IHS' if missing)", type=["xlsx"])
uploaded_parent = st.file_uploader("3. Upload Parent Sites Reference (Parent-sites-Tx...)", type=["xlsx"])

def round_to_nearest_30_minutes(dt):
    """Rounds a datetime object up or down to the nearest 30-minute increment."""
    minute = dt.minute
    remainder = minute % 30
    if remainder < 15:
        dt = dt - timedelta(minutes=remainder)
    else:
        dt = dt + timedelta(minutes=(30 - remainder))
    return dt.replace(second=0, microsecond=0)

# THE TRIGGER NOW ONLY DEMANDS THE ALARM EXPORT AND PARENT SITES TO RUN
if uploaded_alarm and uploaded_parent:
    st.success("Required core files staged successfully.")
    
    if st.button("🚀 Generate Formatted Alarm Report"):
        with st.spinner("Processing network structures and applying styling configurations..."):
            try:
                # 2. INGEST DATA SOURCES
                if uploaded_alarm.name.endswith('.csv'):
                    df_alarm = pd.read_csv(uploaded_alarm)
                else:
                    df_alarm = pd.read_excel(uploaded_alarm)
                    
                df_parent = pd.read_excel(uploaded_parent, sheet_name='Final-Prior-Parent-site-list')

                # Ensure consistent structures and safely convert Site IDs to numeric
                df_alarm['Site ID'] = pd.to_numeric(df_alarm['Site ID'], errors='coerce')
                df_parent.iloc[:, 0] = pd.to_numeric(df_parent.iloc[:, 0], errors='coerce')  # Col A

                # Remove NaN Site IDs to prevent merge issues
                df_alarm = df_alarm.dropna(subset=['Site ID'])
                df_alarm['Site ID'] = df_alarm['Site ID'].astype(int)

                # 3. DATA CLEANING & RE-ORDERING
                keep_cols = ['Site ID', 'Site Name', 'Ticket ID', 'Alarm Name', 'First Occurred On', 'Duration(hh:mm:ss)', 'Last Occurred On']
                df_clean = df_alarm[keep_cols].copy()

                # Prepare parent sites lookup dictionaries
                parent_child_dict = df_parent.dropna(subset=[df_parent.columns[0]]).set_index(df_parent.columns[0])[df_parent.columns[5]].to_dict()
                parent_owner_dict = df_parent.dropna(subset=[df_parent.columns[0]]).set_index(df_parent.columns[0])[df_parent.columns[7]].to_dict()

                # ==========================================
                # DYNAMIC POWER OWNER DEFAULT LOGIC
                # ==========================================
                if uploaded_pm:
                    # If the PM file is uploaded, proceed with normal logic
                    df_pm = pd.read_excel(uploaded_pm, sheet_name='2026_Q2_Planning_MTNC')
                    df_pm.iloc[:, 0] = pd.to_numeric(df_pm.iloc[:, 0], errors='coerce') # Col A
                    pm_dict = df_pm.dropna(subset=[df_pm.columns[0]]).set_index(df_pm.columns[0])[df_pm.columns[4]].to_dict()
                    
                    df_clean['Power Owner'] = df_clean['Site ID'].map(pm_dict)
                    # Fallback to parent file owner if missing in PM file, otherwise fallback to IHS
                    df_clean['Power Owner'] = df_clean['Power Owner'].fillna(df_clean['Site ID'].map(parent_owner_dict)).fillna("IHS")
                else:
                    # If PM file is missing, try parent owner column first, otherwise default everything to IHS
                    st.info("ℹ️ PM Plan file missing. Using parent-site mappings and applying 'IHS' fallback.")
                    df_clean['Power Owner'] = df_clean['Site ID'].map(parent_owner_dict).fillna("IHS")

                # Get internal children metrics for calculation without adding an unnecessary final column
                temp_children = df_clean['Site ID'].map(parent_child_dict)

                # Generate the text predictions directly based on numeric values
                df_clean['Prediction'] = temp_children.apply(
                    lambda x: f"{int(x)} sites will be affected" if pd.notna(x) and isinstance(x, (int, float)) else ""
                )

                # Assemble final requested target column hierarchy
                final_column_order = [
                    'Site ID', 'Site Name', 'Power Owner', 'Prediction', 
                    'Ticket ID', 'Alarm Name', 'First Occurred On', 
                    'Duration(hh:mm:ss)', 'Last Occurred On'
                ]
                df_final = df_clean[final_column_order].copy()

                # Sort globally by Site ID numerical value
                df_final = df_final.sort_values(by='Site ID')

                # 4. EXCEL PRESENTATION ENGINE (openpyxl)
                wb = load_workbook(io.BytesIO())
                ws = wb.active
                ws.title = "Power Alarms Status"
                ws.views.sheetView[0].showGridLines = True

                # Write organized headers and actual data rows
                ws.append(final_column_order)
                for r in dataframe_to_rows(df_final, index=False, header=False):
                    ws.append(r)

                # Define structural aesthetics
                header_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid") # #FFC000 Gold
                header_font = Font(name="Calibri", size=11, bold=True, color="000000")
                data_font = Font(name="Calibri", size=11, color="000000")
                
                thin_side = Side(border_style="thin", color="D3D3D3")
                grid_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

                # Apply layout treatments and add data validation filters
                ws.auto_filter.ref = f"A1:I{ws.max_row}"

                # Style Headers
                for cell in ws[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = grid_border

                # Style Data Grid Rows
                for row in range(2, ws.max_row + 1):
                    for col in range(1, 10):
                        cell = ws.cell(row=row, column=col)
                        cell.font = data_font
                        cell.border = grid_border
                        
                        # Center code indices and dates, left-align names/text
                        if col in [1, 5, 7, 8, 9]:
                            cell.alignment = Alignment(horizontal="center")
                        else:
                            cell.alignment = Alignment(horizontal="left")

                # Auto-adjust tracking widths dynamically
                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = col[0].column_letter
                    ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

                # 5. DYNAMIC TIME STAMP & FILENAME CALCULATION
                current_time = datetime.now()
                rounded_time = round_to_nearest_30_minutes(current_time)
                time_string = rounded_time.strftime("%HH%M")
                filename = f"TIS_ALARMS{time_string}.xlsx"

                # Save resulting spreadsheet asset out to local memory buffers
                output_buffer = io.BytesIO()
                wb.save(output_buffer)
                output_buffer.seek(0)

                st.balloons()
                st.success(f"Report compiled successfully for period context: {time_string}!")
                
                st.download_button(
                    label=f"📥 Download {filename}",
                    data=output_buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            except Exception as e:
                st.error(f"Processing error raised: {str(e)}. Double check that referenced sheet names match exactly.")
else:
    st.warning("Waiting for the Alarm Export and Parent Sites reference file to be uploaded...")
