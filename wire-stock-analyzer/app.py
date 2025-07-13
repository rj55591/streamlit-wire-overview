import pandas as pd
import streamlit as st
import tempfile
import os

def parse_date_fixed(date, customer):
    try:
        if str(customer).strip().lower() == "perniagaan logam hock soon":
            return pd.to_datetime(date, dayfirst=True, errors='coerce')
        return pd.to_datetime(date, errors='coerce')
    except:
        return pd.NaT

def calc_weights(df):
    df = df.rename(columns=lambda x: str(x).strip())
    df["Screen Length"] = pd.to_numeric(df["Screen Length"], errors="coerce")
    df["Screen width"] = pd.to_numeric(df["Screen width"], errors="coerce")
    df["Aperture #"] = pd.to_numeric(df["Aperture #"], errors="coerce")
    df["Wire ø"] = pd.to_numeric(df["Wire ø"], errors="coerce")
    df["QTY"] = pd.to_numeric(df["QTY"], errors="coerce")

    df["Length_mm"] = df["Screen Length"] * 25.4
    df["Width_mm"] = df["Screen width"] * 25.4
    df["Area_mm2"] = df["Length_mm"] * df["Width_mm"]

    df["Weight in kg per item"] = (
        (df["Wire ø"] ** 2) * 12.7 / (df["Aperture #"] + df["Wire ø"])
    ) * (df["Area_mm2"] / 1e6)

    df["Weight in kg total"] = df["Weight in kg per item"] * df["QTY"]
    return df.drop(columns=["Length_mm", "Width_mm", "Area_mm2"])

def generate_final_wire_overview(customer_orders_fp, wire_coil_bal_fp, incoming_stock_fp, default_usage_fp):
    xls = pd.ExcelFile(customer_orders_fp)
    valid_sheets = [s for s in xls.sheet_names if s != "Summary 2025"]
    cleaned_sheets = []

    for sheet in valid_sheets:
        df_raw = xls.parse(sheet, header=None)
        header_row_index = None
        for i in range(len(df_raw)):
            if df_raw.iloc[i].astype(str).str.contains("P/O Date", case=False).any():
                header_row_index = i
                break
        if header_row_index is not None:
            df = pd.read_excel(xls, sheet_name=sheet, header=header_row_index)
            df["Customer"] = sheet
            cleaned_sheets.append(df)

    combined_df = pd.concat(cleaned_sheets, ignore_index=True)
    combined_df.columns = [str(c).strip() for c in combined_df.columns]
    combined_df["P/O Date"] = combined_df.apply(lambda row: parse_date_fixed(row["P/O Date"], row["Customer"]), axis=1)
    filtered_df = combined_df[combined_df["P/O Date"] >= pd.Timestamp("2025-01-01")]

    end_col_index = filtered_df.columns.get_loc("Order No")
    trimmed_df = filtered_df.iloc[:, :end_col_index + 1]
    final_df = calc_weights(trimmed_df)

    pending_df = final_df[final_df["Job Sheet No."].isna()]
    pending_usage = pending_df.groupby("Wire ø")["Weight in kg total"].sum().reset_index()
    pending_usage.columns = ["Wire ø", "Total Pending Wire Required (kg)"]

    coil_xls = pd.ExcelFile(wire_coil_bal_fp)
    wire_inventory_bal = {}
    for sheet in coil_xls.sheet_names:
        try:
            df = coil_xls.parse(sheet, header=4)
            df.columns = [str(c).strip().upper() for c in df.columns]
            if "BAL" in df.columns:
                bal_series = pd.to_numeric(df["BAL"], errors='coerce').dropna()
                if not bal_series.empty:
                    wire_inventory_bal[sheet.strip()] = bal_series.iloc[-1]
        except:
            continue
    inventory_df = pd.DataFrame(list(wire_inventory_bal.items()), columns=["Wire ø", "Available Inventory (kg)"])
    inventory_df["Wire ø"] = pd.to_numeric(inventory_df["Wire ø"], errors="coerce")

    incoming_df = pd.read_excel(incoming_stock_fp)
    incoming_df.columns = [str(c).strip() for c in incoming_df.columns]
    incoming_df = incoming_df.rename(columns={"Wire Diameter": "Wire ø"})
    incoming_df["Wire ø"] = pd.to_numeric(incoming_df["Wire ø"], errors="coerce")
    for col in ["Kewei", "QS", "Bolin"]:
        incoming_df[col] = pd.to_numeric(incoming_df[col], errors="coerce").fillna(0)

    usage_df = pd.read_csv(default_usage_fp)
    usage_df = usage_df.rename(columns={"Avg Jan-May Usage (kg)": "Default Avg Monthly Usage (kg)"})
    usage_df["Wire ø"] = pd.to_numeric(usage_df["Wire ø"], errors="coerce")

    result = pending_usage.merge(inventory_df, on="Wire ø", how="left")
    result = result.merge(incoming_df[["Wire ø", "Kewei", "QS", "Bolin"]], on="Wire ø", how="left")
    result = result.merge(usage_df, on="Wire ø", how="left")
    result.fillna(0, inplace=True)

    return result, pending_df

# Streamlit UI
st.title("Final Wire Overview Generator")

cust_orders_file = st.file_uploader("Upload Customer Order - 2025.05.30.xlsx", type=["xlsx"])
coil_bal_file = st.file_uploader("Upload WIRE COIL BAL (KGS.).xlsx", type=["xlsx"])
incoming_file = st.file_uploader("Upload Incoming Stock.xlsx", type=["xlsx"])
default_usage_file = st.file_uploader("Upload default wire monthly.csv", type=["csv"])

if st.button("Generate Overview") and all([cust_orders_file, coil_bal_file, incoming_file, default_usage_file]):
    cust_fp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    coil_fp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    incoming_fp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    default_fp = tempfile.NamedTemporaryFile(delete=False, suffix=".csv")

    cust_fp.write(cust_orders_file.read())
    coil_fp.write(coil_bal_file.read())
    incoming_fp.write(incoming_file.read())
    default_fp.write(default_usage_file.read())

    cust_fp.close()
    coil_fp.close()
    incoming_fp.close()
    default_fp.close()

    final_df, pending_df = generate_final_wire_overview(
        cust_fp.name,
        coil_fp.name,
        incoming_fp.name,
        default_fp.name
    )

    st.session_state["base_df"] = final_df.copy()
    st.session_state["pending_df"] = pending_df.copy()

if "base_df" in st.session_state:
    base_df = st.session_state["base_df"].copy()

    st.sidebar.header("Supplier Toggle")
    use_kewei = st.sidebar.checkbox("Include Kewei", value=True)
    use_qs = st.sidebar.checkbox("Include QS", value=True)
    use_bolin = st.sidebar.checkbox("Include Bolin", value=True)

    included_suppliers = []
    if use_kewei: included_suppliers.append("Kewei")
    if use_qs: included_suppliers.append("QS")
    if use_bolin: included_suppliers.append("Bolin")

    base_df["Total Available (kg)"] = base_df[["Available Inventory (kg)"] + included_suppliers].sum(axis=1)
    base_df["Surplus / Shortage (kg)"] = base_df["Total Available (kg)"] - base_df["Total Pending Wire Required (kg)"]
    base_df["Months of Coverage"] = (
        base_df["Surplus / Shortage (kg)"] / base_df["Default Avg Monthly Usage (kg)"].replace({0: pd.NA})
    ).round(2)

    st.success("Overview generated successfully!")
    included_text = ", ".join(included_suppliers) if included_suppliers else "None"
    st.markdown(f"**ℹ️ Calculations include suppliers:** `{included_text}`")
    st.dataframe(base_df)

    export_fp = os.path.join(tempfile.gettempdir(), "Final_Wire_Overview.xlsx")
    base_df.to_excel(export_fp, index=False)
    with open(export_fp, "rb") as f:
        st.download_button(
            label="Download Final Wire Overview as XLSX",
            data=f,
            file_name="Final_Wire_Overview.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    pending_fp = os.path.join(tempfile.gettempdir(), "Pending_Orders.xlsx")
    st.session_state["pending_df"].to_excel(pending_fp, index=False)
    with open(pending_fp, "rb") as f:
        st.download_button(
            label="Download Pending Orders as XLSX",
            data=f,
            file_name="Pending_Orders.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
