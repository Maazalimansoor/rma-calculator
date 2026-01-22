import pandas as pd

# Load Excel sheet
df_items = pd.read_excel("RMA calculator.xlsx", sheet_name=0, header=0)

# Strip leading/trailing spaces from column names
df_items.columns = df_items.columns.str.strip()

# Print columns to debug
print("Columns in Items sheet:", df_items.columns.tolist())

# Find the correct column for Item
item_col_candidates = [col for col in df_items.columns if 'item' in col.lower()]
if not item_col_candidates:
    raise ValueError("No column found for Item in Items sheet")
item_col = item_col_candidates[0]  # use the first matching column
print("Using column for Item:", item_col)
