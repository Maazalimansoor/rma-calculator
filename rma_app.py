import streamlit as st
import pandas as pd
import os

# ------------------ PAGE CONFIG ------------------
st.set_page_config(page_title="RMA Calculator", layout="wide")

# Hide header/footer
st.markdown(
    """
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .block-container {padding-top: 0rem; padding-left: 1rem; padding-right: 1rem;}
    </style>
    """,
    unsafe_allow_html=True
)

st.title("**Welcome To Alpha RMA Cost Calculator**")
st.write("NEW VERSION LIVE")
# ------------------ CONSTANTS ------------------
OUTBOUND_SHIPPING = 5
RETURN_SHIPPING = 10
REPLACEMENT_SHIPPING = 15
ADMIN_COST = 20
QC_COST = 5
REPAIR_MATERIAL = 5
REPAIR_LABOR = 35

# ------------------ SIDEBAR ------------------
st.sidebar.header("Adjust Shipping Costs (Optional)")
return_shipping = st.sidebar.number_input("Return Shipping", value=RETURN_SHIPPING)
outbound_shipping = st.sidebar.number_input("Outbound Shipping", value=OUTBOUND_SHIPPING)
replacement_shipping = st.sidebar.number_input("Replacement Shipping", value=REPLACEMENT_SHIPPING)
st.sidebar.header("QC Cost : $5")
st.sidebar.header("Admin Cost : $20")

# ------------------ LOAD DATA ------------------
@st.cache_data
def load_data():
    file_path = os.path.join(os.path.dirname(__file__), "RMA calculator.xlsx")
    df_items = pd.read_excel(file_path, sheet_name=0)
    df_dropdowns = pd.read_excel(file_path, sheet_name=2)
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
item_list = ["Select an Item"] + df_items["item"].astype(str).unique().tolist()
item_input = st.selectbox("Enter Item Number", item_list, index=0)
if item_input == "Select an Item":
    st.stop()

qty = st.number_input("Enter Quantity", min_value=1, value=1)

# ------------------ CALCULATE COSTS ------------------
def calculate_rma(item_input, qty):
    row = df_items[df_items["item"].astype(str) == str(item_input)].iloc[0]
    selling_price = float(row["regprice"])
    std_cost = float(row["stdcost"])

    rma_types = [
        "Credit with Return (Good product)",
        "Credit Only (No Return, Scrap at Customer)",
        "Credit with Return (scrap on the ALPHA site)",
        "Replacement (if return product is bad)",
        "Replacement (Return Product is good & resellable)",
        "Scrap after return",
        "Repair (IDF / RMA)",
        "Ship labels only",
        "Ship Box and Labels Only",
        "Non-Warranty Claim",
        "Missing Shipment"
    ]

    rma_results = []
    for rma in rma_types:
        if rma == "Credit with Return (Good product)":
            cost = ((selling_price + return_shipping + outbound_shipping + QC_COST) * qty + ADMIN_COST)
        elif rma == "Credit Only (No Return, Scrap at Customer)":
            cost = ((selling_price + std_cost) * qty + ADMIN_COST)
        elif rma == "Scrap after return":
            cost = ((std_cost + return_shipping + outbound_shipping + QC_COST) * qty + ADMIN_COST)
        elif rma == "Replacement (if return product is bad)":
            cost = ((std_cost + outbound_shipping + QC_COST*2 + std_cost + replacement_shipping) * qty + ADMIN_COST)
        elif rma == "Repair (IDF / RMA)":
            cost = ((REPAIR_MATERIAL + REPAIR_LABOR + QC_COST + replacement_shipping) * qty + ADMIN_COST)
        elif rma == "Credit with Return (scrap on the ALPHA site)":
            cost = ((selling_price + return_shipping + outbound_shipping + QC_COST + std_cost) * qty + ADMIN_COST)
        elif rma == "Ship labels only":
            cost = ((outbound_shipping + QC_COST + 1) * qty + ADMIN_COST)
        elif rma == "Ship Box and Labels Only":
            cost = ((outbound_shipping + QC_COST + 2) * qty + ADMIN_COST)
        elif rma == "Replacement (Return Product is good & resellable)":
            cost = ((outbound_shipping + QC_COST*2 + replacement_shipping) * qty + ADMIN_COST)
        elif rma == "Non-Warranty Claim":
            cost = ((outbound_shipping + QC_COST + REPAIR_LABOR + REPAIR_MATERIAL + replacement_shipping) * qty + ADMIN_COST)
        elif rma == "Missing Shipment":
            cost = ((QC_COST + std_cost + outbound_shipping) * qty + ADMIN_COST)
        else:
            cost = 0

        # Categorize
        if "Credit" in rma:
            category = "Credit"
        elif "Replacement" in rma:
            category = "Replacement"
        elif "Repair" in rma:
            category = "Repair"
        else:
            category = "Other"

        rma_results.append({"RMA Type": rma, "Total Cost": cost, "Category": category})

    return pd.DataFrame(rma_results), selling_price, std_cost

# Calculate dynamically every run
df_results, selling_price, std_cost = calculate_rma(item_input, qty)

# ------------------ SHOW ITEM INFO ------------------
st.markdown(f"**Item:** {item_input}")
col1, col2 = st.columns(2)
col1.metric("STD Cost", f"${std_cost:,.2f}")
col2.metric("Regular Price", f"${selling_price:,.2f}")

# ------------------ CATEGORY FILTER BUTTONS ------------------
if "selected_category" not in st.session_state:
    st.session_state.selected_category = "All"

st.markdown("### Filter by RMA Category")
categories = ["All"] + df_results["Category"].unique().tolist()
cols = st.columns(len(categories))

for i, cat in enumerate(categories):
    if cols[i].button(cat):
        st.session_state.selected_category = cat

# Filter table
if st.session_state.selected_category != "All":
    df_filtered = df_results[df_results["Category"] == st.session_state.selected_category]
else:
    df_filtered = df_results.copy()

# ------------------ STYLE TABLE ------------------
def style_rma_table(row):
    rma = row["RMA Type"]
    total = row["Total Cost"]

    if "Credit" in rma:
        color = "lightblue"
    elif "Replacement" in rma:
        color = "lightsalmon"
    elif "Repair" in rma:
        color = "plum"
    else:
        color = "lightgray"

    if total == df_filtered["Total Cost"].min():
        color = "lightgreen"
    elif total == df_filtered["Total Cost"].max():
        color = "yellow"

    return ["background-color: {}".format(color) for _ in row]

df_styled = df_filtered.style.format({"Total Cost": "${:,.2f}"}).apply(style_rma_table, axis=1)

# ------------------ DISPLAY ------------------
st.success("RMA Costs Calculated")

# Legend
st.markdown(
    """
    <b>Legend:</b><br>
    <span style='background-color: lightblue'>&nbsp;&nbsp;&nbsp;&nbsp;</span> Credit &nbsp;&nbsp;
    <span style='background-color: lightsalmon'>&nbsp;&nbsp;&nbsp;&nbsp;</span> Replacement &nbsp;&nbsp;
    <span style='background-color: plum'>&nbsp;&nbsp;&nbsp;&nbsp;</span> Repair &nbsp;&nbsp;
    <span style='background-color: lightgray'>&nbsp;&nbsp;&nbsp;&nbsp;</span> Other &nbsp;&nbsp;
    <span style='background-color: lightgreen'>&nbsp;&nbsp;&nbsp;&nbsp;</span> Min Cost &nbsp;&nbsp;
    <span style='background-color: yellow'>&nbsp;&nbsp;&nbsp;&nbsp;</span> Max Cost
    """,
    unsafe_allow_html=True
)

# Min/max summary
if not df_filtered.empty:
    min_cost = df_filtered["Total Cost"].min()
    max_cost = df_filtered["Total Cost"].max()
    st.markdown(f"**Minimum RMA Cost:** ${min_cost:,.2f}  |  **Maximum RMA Cost:** ${max_cost:,.2f}")

# Table height
table_height = max(len(df_filtered)*40, 200)
st.dataframe(df_styled, use_container_width=False, height=table_height)
