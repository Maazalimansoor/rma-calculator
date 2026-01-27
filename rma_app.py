import streamlit as st
import pandas as pd
import os

# ================== PAGE CONFIG ==================
st.set_page_config(page_title="Alpha RMA Cost Calculator", layout="wide")

st.markdown("""
<style>
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
.block-container {padding-top: 0.5rem;}
</style>
""", unsafe_allow_html=True)

st.title("Alpha RMA Cost Calculator")

# ================== CONSTANTS ==================
ADMIN_COST = 20
QC_COST = 5
REPAIR_MATERIAL = 5
REPAIR_LABOR = 35

DEFAULT_OUTBOUND = 0
DEFAULT_RETURN = 0
DEFAULT_REPLACEMENT = 0

# ================== SIDEBAR ==================
st.sidebar.header("Shipping Costs $ per electrode")
outbound_shipping = st.sidebar.number_input("Outbound Shipping $", value=DEFAULT_OUTBOUND)
return_shipping = st.sidebar.number_input("Return Shipping $", value=DEFAULT_RETURN)
replacement_shipping = st.sidebar.number_input("Replacement Shipping $", value=DEFAULT_REPLACEMENT)

QC_COST = st.sidebar.number_input("QC Cost per electrode $", value=5.0, step=0.1, format="%.2f")
ADMIN_COST = st.sidebar.number_input("Admin Cost per event $", value=20.0, step=0.1, format="%.2f")

st.sidebar.header("Calculations $")
st.sidebar.write("**Credit for Return of Functional Product:** = (Price + Return shipping + Outbound shipping + QC) × Qty + Admin Cost")
st.sidebar.write("**Credit Only – Product Scrapped at Customer:** = (Price × Qty) + Admin Cost")
st.sidebar.write("**Credit Issued – Return Received, Scrapped at ALPHA:** = (Price + Return shipping + Outbound shipping + QC + Std Cost) × Qty + Admin Cost")
st.sidebar.write("**Replacement Provided Only for Non-Functional Returned Product:** = (Std Cost + Outbound + Replacement + 2×QC + Std Cost) × Qty + Admin Cost")
st.sidebar.write("**Replacement Issued – Returned Product Good & Resellable:** = (Outbound + Replacement + 2×QC) × Qty + Admin Cost")
st.sidebar.write("**Returned for Repair (IDF / RMA):** = (Repair Material + Labor + QC + Replacement) × Qty + Admin Cost")
st.sidebar.write("**Shipping Labels Only:** = (Outbound + QC + $1) × Qty + Admin Cost")
st.sidebar.write("**Shipping Box and Labels Only:** = (Outbound + QC + $2) × Qty + Admin Cost")
st.sidebar.write("**Non-Warranty Claim:** = (Repair Material + Labor + Outbound + Replacement + QC) × Qty + Admin Cost")
st.sidebar.write("**Shipment Not Received – Replacement Processed:** = (Std Cost + Outbound + QC) × Qty + Admin Cost")

# ================== LOAD EXCEL ==================
@st.cache_data(show_spinner="Loading Excel data...")
def load_excel():
    file_path = os.path.join(os.path.dirname(__file__), "RMA calculator.xlsx")
    if not os.path.exists(file_path):
        st.error("Excel file not found")
        st.stop()

    df = pd.read_excel(file_path, sheet_name=0)
    df.columns = df.columns.str.strip().str.lower().str.replace(" ", "")
    return df

if st.button("🔄 Refresh Excel Data"):
    st.cache_data.clear()

df_items = load_excel()

# ================== VALIDATION ==================
required_cols = ["item", "regprice", "stdcost"]
missing = [c for c in required_cols if c not in df_items.columns]
if missing:
    st.error(f"Missing required columns: {missing}")
    st.stop()

# ================== USER INPUT ==================
item_list = ["Select an Item"] + sorted(df_items["item"].astype(str).unique().tolist())
item_input = st.selectbox("Enter Part Number", item_list)

if item_input == "Select an Item":
    st.stop()

qty = st.number_input("Enter Quantity", min_value=0.01, value=1.00, step=0.01)

row = df_items[df_items["item"].astype(str) == str(item_input)].iloc[0]
selling_price = float(row["regprice"])
std_cost = float(row["stdcost"])
misc10_value = row.get("misc10", "")
onhand_value = row.get("onhand", "")

# ================== MARGIN ==================
margin_percent = ((selling_price - std_cost) / selling_price) * 100 if selling_price else 0
if margin_percent < 30: margin_color = "#FF6347"
elif margin_percent < 50: margin_color = "#FFA500"
elif margin_percent < 70: margin_color = "#FFFF99"
elif margin_percent < 90: margin_color = "#90EE90"
else: margin_color = "#32CD32"

# ================== KPI DISPLAY ==================
st.markdown("### Item Information")
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric("Standard Cost", f"${std_cost:,.2f}")
c2.metric("Regular Price", f"${selling_price:,.2f}")
c3.metric("Lab-Ovh", misc10_value)
c4.metric("On-Hand", onhand_value)
c5.markdown(f"""<div style="text-align:center;"><div style="font-weight:bold;">Margin %</div>
<div style="background:{margin_color};padding:14px;border-radius:10px;font-size:20px;font-weight:bold;">{margin_percent:.2f}%</div></div>""", unsafe_allow_html=True)

# ================== LEGEND ==================
st.markdown("##### Margin Legend")
l1,l2,l3,l4,l5 = st.columns(5)
l1.markdown('<div style="background:#FF6347;height:12px;border-radius:4px;"></div><small>&lt;30%</small>', unsafe_allow_html=True)
l2.markdown('<div style="background:#FFA500;height:12px;border-radius:4px;"></div><small>30–50%</small>', unsafe_allow_html=True)
l3.markdown('<div style="background:#FFFF99;height:12px;border-radius:4px;"></div><small>50–70%</small>', unsafe_allow_html=True)
l4.markdown('<div style="background:#90EE90;height:12px;border-radius:4px;"></div><small>70–90%</small>', unsafe_allow_html=True)
l5.markdown('<div style="background:#32CD32;height:12px;border-radius:4px;"></div><small>≥90%</small>', unsafe_allow_html=True)

# ================== RMA CALCULATIONS ==================
rma_types = [
    "Credit for Return of Functional Product",
    "Credit Only – Product Scrapped at Customer",
    "Credit Issued – Return Received, Scrapped at ALPHA",
    "Replacement Provided Only for Non-Functional Returned Product",
    "Replacement Issued – Returned Product Good & Resellable",
    "Returned for Repair (IDF / RMA)",
    "Shipping Labels Only",
    "Shipping Box and Labels Only",
    "Non-Warranty Claim",
    "Shipment Not Received – Replacement Processed",
]

results = []
# ================== RMA CALCULATIONS ==================

for rma in rma_types:
    if rma == "Credit for Return of Functional Product":
        cost = (selling_price + return_shipping + outbound_shipping + QC_COST) * qty + ADMIN_COST
        category = "Credit"
    elif rma == "Credit Only – Product Scrapped at Customer":
        cost = (selling_price * qty) + ADMIN_COST
        category = "Credit"
    elif rma == "Credit Issued – Return Received, Scrapped at ALPHA":
        cost = (selling_price + return_shipping + outbound_shipping + QC_COST + std_cost) * qty + ADMIN_COST
        category = "Credit"
    elif rma == "Replacement Provided Only for Non-Functional Returned Product":
        cost = (std_cost + outbound_shipping + replacement_shipping + 2 * QC_COST + std_cost) * qty + ADMIN_COST
        category = "Replacement"
    elif rma == "Replacement Issued – Returned Product Good & Resellable":
        cost = (outbound_shipping + replacement_shipping + 2 * QC_COST) * qty + ADMIN_COST
        category = "Replacement"
    elif rma == "Returned for Repair (IDF / RMA)":
        cost = (REPAIR_MATERIAL + REPAIR_LABOR + QC_COST + replacement_shipping) * qty + ADMIN_COST
        category = "Repair"
    elif rma == "Shipping Labels Only":
        cost = (outbound_shipping + QC_COST + 1) * qty + ADMIN_COST
        category = "Other"
    elif rma == "Shipping Box and Labels Only":
        cost = (outbound_shipping + QC_COST + 2) * qty + ADMIN_COST
        category = "Other"
    elif rma == "Non-Warranty Claim":
        cost = (REPAIR_MATERIAL + REPAIR_LABOR + outbound_shipping + replacement_shipping + QC_COST) * qty + ADMIN_COST
        category = "Other"
    elif rma == "Shipment Not Received – Replacement Processed":
        cost = (std_cost + outbound_shipping + QC_COST) * qty + ADMIN_COST
        category = "Other"
    else:
        cost = 0
        category = "Other"

    results.append({"RMA Type": rma, "Category": category, "Total Cost": cost})

df_results = pd.DataFrame(results)


# ================== FILTER BUTTONS ==================
if "category" not in st.session_state:
    st.session_state.category = "All"

st.markdown("### Filter by RMA Category")
cats = ["All"] + sorted(df_results["Category"].unique().tolist())
cols = st.columns(len(cats))

for i, c in enumerate(cats):
    if cols[i].button(c):
        st.session_state.category = c

# Apply the filter
if st.session_state.category != "All":
    df_view = df_results[df_results["Category"] == st.session_state.category]
else:
    df_view = df_results.copy()

# ================== STYLE ==================
def style_row(row):
    if row["Total Cost"] == df_results["Total Cost"].min():
        color = "#D5F5E3"
    elif row["Total Cost"] == df_results["Total Cost"].max():
        color = "#FCF3CF"
    elif row["Category"] == "Credit":
        color = "#D6EAF8"
    elif row["Category"] == "Replacement":
        color = "#FADBD8"
    elif row["Category"] == "Repair":
        color = "#E8DAEF"
    else:
        color = "#EAECEE"
    return [f"background-color:{color}; font-weight:bold;" for _ in row]

df_styled = df_view.style.format({"Total Cost": "${:,.2f}"}).apply(style_row, axis=1)

# ================== DISPLAY ==================
st.success("RMA Costs Calculated")
st.dataframe(df_styled, use_container_width=False, height=450)

min_cost = df_view["Total Cost"].min()
max_cost = df_view["Total Cost"].max()

col_min, col_max = st.columns(2)
col_min.metric("Minimum RMA Cost", f"${min_cost:,.2f}")
col_max.metric("Maximum RMA Cost", f"${max_cost:,.2f}")