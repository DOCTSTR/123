import streamlit as st
import pandas as pd
import os
import tempfile
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font
from zipfile import ZipFile

# Title and Developer Name
st.set_page_config(page_title="SID-FIR Dashboard", layout="wide")
st.title("📂 SID ↔ FIR Matching Tool")
st.markdown("### Developed By **Doctstr (Meghraj Police Station)**")

# Sidebar Dropdown
mode = st.selectbox("Select Mode", ["FIR Link in SID", "SID Use for FIR"])

# File Uploads
fir_file = st.file_uploader("Upload FIR Excel File (.xls)", type=["xls"])
sid_files = st.file_uploader("Upload Multiple SID Excel Files (.xls)", type=["xls"], accept_multiple_files=True)

# Help Button
with st.sidebar:
    st.markdown("### ❓ Help Section")
    help_file = st.file_uploader("Upload Help PDF", type="pdf")
    if help_file:
        st.download_button("📖 View Help PDF", help_file, file_name="Help_Document.pdf")

# Function to handle the logic of both modes
def process_files(fir_file, sid_files, mode):
    # Police Station Mapping
    police_station_mapping = {
        "11188003": "ભીલોડા", "11188010": "શામળાજી", "11188004": "ધનસુરા",
        "11188002": "બાયડ", "11188001": "આબલીયારા", "11188009": "મોડાસા_ટાઉન",
        "11188008": "મોડાસા_રૂરલ", "11188007": "મેધરજ", "11188006": "માલપુર",
        "11188005": "ઇસરી", "11188011": "સાથંબા", "11188012": "મહિલા_પોલીસ_સ્ટેશન",
        "11188013": "ટીંટોઇ", "11188014": "સાયબર ક્રાઇમ પોલીસ સ્ટેશન",
    }

    # Read SID files
    sid_df_list = []
    for file in sid_files:
        df = pd.read_excel(file, engine='xlrd', header=None)
        sid_df_list.append(df)
    merged_sid_df = pd.concat(sid_df_list, ignore_index=True)

    # Read FIR file
    df2 = pd.read_excel(fir_file, engine='xlrd', header=None)
    police_station_name = df2.iloc[4, 1]
    date_column = df2.iloc[4:, 2].dropna()
    start_date = pd.to_datetime(date_column.iloc[0], dayfirst=True).strftime("%d/%m/%Y")
    end_date = pd.to_datetime(date_column.iloc[-1], dayfirst=True).strftime("%d/%m/%Y")

    # Extract FIR Numbers
    fir_number = df2.iloc[4:, 1].reset_index(drop=True)

    # Case Data Extraction
    case_number_1 = merged_sid_df.iloc[3:, 2].reset_index(drop=True)
    case_number_2 = merged_sid_df.iloc[3:, 10].reset_index(drop=True)

    if mode == "FIR Link in SID":
        output_df = pd.DataFrame({"Case Number 2": case_number_2, "FIR Number": fir_number})
        output_df["Final Output"] = output_df["FIR Number"].apply(lambda x: x if x in case_number_2.values else None)
    else:
        output_df = pd.DataFrame({
            "Case_Number_1": case_number_1,
            "Case Number 2": case_number_2,
            "FIR Number": fir_number
        })
        all_cases = pd.concat([case_number_1, case_number_2]).dropna().unique()
        output_df["Final Output"] = output_df["FIR Number"].apply(lambda x: x if x in all_cases else None)

    output_df["Pending SID"] = output_df.apply(
        lambda row: row["FIR Number"] if pd.isna(row["Final Output"]) else None, axis=1
    )

    # Summary Metrics
    fir_filled_count = output_df["FIR Number"].count()
    final_filled_count = output_df["Final Output"].count()
    pending_sid_count = output_df["Pending SID"].count()

    # FIR Prefix and Station
    output_df["FIR Prefix"] = output_df["FIR Number"].astype(str).str[:8]
    output_df["Mapped Police Station"] = output_df["FIR Prefix"].map(police_station_mapping)

    # Sheet 2
    output_df_sorted = output_df.sort_values(by=["FIR Prefix", "FIR Number"])
    io_map = dict(zip(df2.iloc[4:, 1], df2.iloc[4:, 6]))
    sheet2_data, last_prefix = [], None

    for _, row in output_df_sorted.iterrows():
        fir_prefix = row["FIR Prefix"]
        station = row["Mapped Police Station"]
        fir_num = row["FIR Number"]
        final_out = row["Final Output"]
        pending = row["Pending SID"]
        pending_link = pending if pd.notna(pending) else None
        io_name = io_map.get(pending_link, "") if pending_link else ""
        sheet2_data.append([
            fir_prefix if fir_prefix != last_prefix else '',
            station if fir_prefix != last_prefix else '',
            fir_num, final_out, pending, pending_link, io_name
        ])
        last_prefix = fir_prefix

    sheet2_df = pd.DataFrame(sheet2_data, columns=[
        "FIR Prefix", "Mapped Police Station", "FIR Number", "Final Output",
        "Pending SID", "Pending Fir Link", "IO Name"
    ])

    # Sheet 3 - Dashboard
    dashboard_data = []
    for station in output_df["Mapped Police Station"].dropna().unique():
        group = output_df[output_df["Mapped Police Station"] == station]
        fir_count = group["FIR Number"].count()
        final_count = group["Final Output"].count()
        pending_count = group["Pending SID"].count()
        percentage = round((final_count / fir_count) * 100, 2) if fir_count else 0
        dashboard_data.append([station, fir_count, final_count, pending_count, percentage])

    dashboard_df = pd.DataFrame(
        dashboard_data,
        columns=["પો.સ્ટેનુ નામ", "એફ.આઇ.આર સંખ્યા", "SID સંખ્યા", "SID બાકી સંખ્યા", "ટકાવારી"]
    )
    dashboard_df = dashboard_df.sort_values(by="ટકાવારી", ascending=False).reset_index(drop=True)
    dashboard_df.insert(0, "ક્રમ સં.", range(1, len(dashboard_df) + 1))
    title_row = pd.DataFrame([[f"E-Sakshya SID  Dt.{start_date} To Dt.{end_date}", None, None, None, None, None]],
                             columns=dashboard_df.columns)
    header_row = pd.DataFrame([dashboard_df.columns.tolist()], columns=dashboard_df.columns)
    total_row = pd.DataFrame([[ "", "કુલ",
        dashboard_df["એફ.આઇ.આર સંખ્યા"].sum(),
        dashboard_df["SID સંખ્યા"].sum(),
        dashboard_df["SID બાકી સંખ્યા"].sum(),
        round((dashboard_df["SID સંખ્યા"].sum() / dashboard_df["એફ.આઇ.આર સંખ્યા"].sum()) * 100, 2)
    ]], columns=dashboard_df.columns)
    sheet3_df = pd.concat([title_row, header_row, dashboard_df, total_row], ignore_index=True)

    # Save Excel to BytesIO
    excel_output = BytesIO()
    with pd.ExcelWriter(excel_output, engine='openpyxl') as writer:
        output_df.to_excel(writer, index=False, sheet_name="Sheet1")
        sheet2_df.to_excel(writer, index=False, sheet_name="Sheet2")
        sheet3_df.to_excel(writer, index=False, header=False, sheet_name="Sheet3")
        writer.book["Sheet3"]["A2"].font = Font(bold=True)
        for cell in writer.book["Sheet3"][writer.book["Sheet3"].max_row]:
            cell.font = Font(bold=True)

    excel_output.seek(0)
    return excel_output, sheet3_df

# Run Processing
if st.button("🔍 Generate Output") and fir_file and sid_files:
    with st.spinner("Processing... Please wait."):
        excel_bytes, sheet3_output = process_files(fir_file, sid_files, mode)

        # Display Sheet3
        st.subheader("📊 Dashboard Preview (Sheet3)")
        st.dataframe(sheet3_output)

        # Download Button
        st.download_button("📥 Download Output Excel", data=excel_bytes, file_name="Output.xlsx")

