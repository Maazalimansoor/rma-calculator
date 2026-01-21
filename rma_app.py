# force rebuild
import streamlit as st
import pandas as pd
import os

st.set_page_config(page_title="RMA Calculator", layout="wide")
st.title("Welcome To RMA Calculator - Full RMA Cost Calculator")

# ------------------ CONSTANTS ------------------
OUTBOUND_SHIPPING = 5
RETURN_SHIPPING = 10
REPLACEMENT_SHIPPING = 15
ADMIN_COST = 20
QC_COST = 10.50
REPAIR_MATERIAL = 5
REPAIR_LABOR = 35

# ------------------ LOAD DATA (CACHED) ------------------
@st.cache_data
def load_data():
    file_path = os.path.join(os.path.dirname(__file__), "RMA calculator.xlsx")
    df_items = pd.read_excel(file_path, sheet_name=0, engine="openpyxl")
    df_dropdowns = pd.read_excel(file_path, sheet_name=2, engine="openpyxl")
    df_items.columns = df_items.columns.str.strip().str.lower()
    df_dropdowns.columns = df_dropdowns.columns.str.strip()
    return df_items, df_dropdowns

df_items, df_dropdowns = load_data()

# ------------------ CHECK REQUIRED COLUMNS ------------------
required_cols = ["item", "regprice", "stdcost"]
missing = [c for c in required_cols if c not in df_items.columns]
if missing:
    st.error(f"Missing required columns in All Immaster sheet: {missing}")
    st.stop()

# ------------------ USER INPUTS ------------------
item_input = st.selectbox(
    "Select Item",
    df_items["item"].astype(str).unique()
)

qty = st.number_input("Quantity", min_value=1, value=1)

calculate = st.button("Calculate RMA Cost")
if not calculate:
    st.stop()

# ------------------ FETCH ITEM DATA ------------------
row = df_items[df_items["item"].astype(str) == str(item_input)].iloc[0]
selling_price = float(row["regprice"])
std_cost = float(row["stdcost"])

# ------------------ SHOW ITEM INFO ------------------
st.markdown(f"**Item:** {item_input}")
st.markdown(f"**STD Cost:** ${std_cost:,.2f} | **RegPrice:** ${selling_price:,.2f}")
st.markdown("---")

# ------------------ RMA TYPES ------------------
rma_types = [
    "Credit with Return (Good product)",
    "Credit Only (No Return, Scrap at Customer)",
    "If Scrap after return",
    "Replacement (if return product is bad)",
    "Repair (IDF / RMA)",
    "If Credit with (scrap on the ALPHA site)",
    "Ship labels only",
    "Ship Box and Labels Only",
    "REPLACEMENT (RETURNED PRODUCT IS GOOD & RESELLABLE)",
    "Non-Warranty Claim",
    "Missing Shipment"
]

# ------------------ CALCULATE ALL RMA COSTS ------------------
rma_results = []

for rma in rma_types:
    if rma == "Credit with Return (Good product)":
        cost = ((selling_price + RETURN_SHIPPING + OUTBOUND_SHIPPING + QC_COST) * qty + ADMIN_COST)
    elif rma == "Credit Only (No Return, Scrap at Customer)":
        cost = ((selling_price + std_cost )* qty + ADMIN_COST)
    elif rma == "If Scrap after return":
        cost = ((std_cost + RETURN_SHIPPING + OUTBOUND_SHIPPING + QC_COST)* qty + ADMIN_COST)
    elif rma == "Replacement (if return product is bad)":
        cost = ((std_cost + OUTBOUND_SHIPPING + QC_COST + QC_COST + std_cost + REPLACEMENT_SHIPPING) * qty + ADMIN_COST)
    elif rma == "Repair (IDF / RMA)":
        cost = ((REPAIR_MATERIAL + REPAIR_LABOR + QC_COST + REPLACEMENT_SHIPPING) * qty) + ADMIN_COST
    elif rma == "If Credit with (scrap on the ALPHA site)":
        cost = ((selling_price + RETURN_SHIPPING + OUTBOUND_SHIPPING + QC_COST + std_cost) * qty + ADMIN_COST)
    elif rma == "Ship labels only":
        cost = ((OUTBOUND_SHIPPING + QC_COST + 1) * qty + ADMIN_COST)
    elif rma == "Ship Box and Labels Only":
        cost = ((OUTBOUND_SHIPPING + QC_COST + 1+1) * qty + ADMIN_COST)
    elif rma == "REPLACEMENT (RETURNED PRODUCT IS GOOD & RESELLABLE)":
        cost = ((OUTBOUND_SHIPPING + QC_COST + QC_COST + REPLACEMENT_SHIPPING) * qty + ADMIN_COST)
    elif rma == "Non-Warranty Claim":
        cost = ((OUTBOUND_SHIPPING + QC_COST + REPAIR_LABOR + REPAIR_MATERIAL + REPLACEMENT_SHIPPING) * qty + ADMIN_COST)
    elif rma == "Missing Shipment":
        cost = ((QC_COST + std_cost + OUTBOUND_SHIPPING) * qty + ADMIN_COST)
    else:
        cost = 0

    rma_results.append({
        "RMA Type": rma,
        "Total Cost": cost
    })

# ------------------ CREATE DATAFRAME ------------------
df_results = pd.DataFrame(rma_results)

# ------------------ HIGHLIGHT MIN/MAX ------------------
def highlight_extremes(s):
    is_max = s == s.max()
    is_min = s == s.min()
    return ['background-color: red' if v else 'background-color: lightgreen' if m else '' for v, m in zip(is_max, is_min)]

df_styled = df_results.style.format({"Total Cost": "${:,.2f}"}).apply(highlight_extremes, subset=["Total Cost"])

# ------------------ DISPLAY ------------------
st.success("RMA Costs Calculated")
# Use dataframe for compact display
st.dataframe(df_styled, use_container_width=True)


