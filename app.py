import os
import openpyxl
import pandas as pd
from PIL import Image
import streamlit as st

# --- Page Configuration ---
st.set_page_config(
    page_title="IPA-1 Wellness Inventory Control", page_icon="💊", layout="wide"
)

EXCEL_FILE = "Active supply list.xlsx"
LOGO_FILE = "image.png"
SHEET_NAME = "United States"


# --- Load Data Function ---
@st.cache_data(ttl=1)
def load_data():
  if not os.path.exists(EXCEL_FILE):
    st.error(f"Excel file '{EXCEL_FILE}' not found!")
    return pd.DataFrame()

  df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET_NAME)
  # Ensure expected columns exist
  expected_cols = [
      "Item Description",
      "Coupa UIN",
      "PAR Level",
      "Current Stock",
      "Variance ",
  ]
  return df


# --- Header & Branding (Requirement 4) ---
col_logo, col_title = st.widgets.columns([1, 4]) if hasattr(st, "widgets") else st.columns([1, 4])

with col_logo:
  if os.path.exists(LOGO_FILE):
    logo = Image.open(LOGO_FILE)
    st.image(logo, width=120)
  else:
    st.warning("Logo image.png not found.")

with col_title:
  st.title("IPA-1 Wellness Inventory Control")
  st.markdown(
      "*Onsite Medical Representative (WHS) — Inventory & Reorder Management System*"
  )

st.markdown("---")

# Load current data
df_data = load_data()

if df_data.empty:
  st.stop()

# --- Sidebar Navigation ---
st.sidebar.title("Navigation")
app_mode = st.sidebar.radio(
    "Select View",
    [
        "📦 Full Inventory & Master List",
        "📝 Monthly Inventory Entry & Update",
        "🛒 Order Sheet (Below Par)",
        "➕ Add New Item",
    ],
)

# Normalize column names for internal handling if needed
# Expected columns: ['Item Description', 'Coupa UIN', 'PAR Level', 'Current Stock', 'Variance ']


# ==========================================
# 1. FULL INVENTORY & MASTER LIST
# ==========================================
if app_mode == "📦 Full Inventory & Master List":
  st.subheader("Master Supply Inventory Database")
  st.markdown(
      "View and search all active medical and safety supplies linked directly"
      " to your Excel dataset."
  )

  search_query = st.text_input(
      "Search by Description or Coupa UIN", ""
  )

  display_df = df_data.copy()
  if search_query:
    display_df = display_df[
        display_df["Item Description"]
        .astype(str)
        .str.contains(search_query, case=False, na=False)
        | display_df["Coupa UIN"]
        .astype(str)
        .str.contains(search_query, case=False, na=False)
    ]

  st.dataframe(display_df, use_container_width=True, height=500)
  st.info(f"Total Items Tracked: {len(df_data)}")


# ==========================================
# 2. MONTHLY INVENTORY ENTRY & UPDATE
# ==========================================
elif app_mode == "📝 Monthly Inventory Entry & Update":
  st.subheader("Monthly Inventory Count Entry")
  st.markdown(
      "Enter your physical inventory counts below. Changes update the stock list"
      " in real-time and save directly to the Excel document."
  )

  with st.form("inventory_update_form"):
    updated_df = df_data.copy()

    # Use data_editor for intuitive inline spreadsheet updating
    edited_df = st.data_editor(
        updated_df,
        column_config={
            "Item Description": st.column_config.TextColumn(disabled=True),
            "Coupa UIN": st.column_config.TextColumn(disabled=True),
            "PAR Level": st.column_config.NumberColumn(
                "PAR Level", min_value=0, step=1
            ),
            "Current Stock": st.column_config.NumberColumn(
                "Current Stock", min_value=0, step=1
            ),
            "Variance ": st.column_config.NumberColumn(disabled=True),
        },
        use_container_width=True,
        height=550,
        key="inventory_editor",
    )

    submitted = st.form_submit_button("Save & Update Excel Document")

    if submitted:
      # Recalculate variance dynamically: Current Stock - PAR Level (matching Excel formula =D2-C2)
      edited_df["Variance "] = (
          edited_df["Current Stock"] - edited_df["PAR Level"]
      )

      try:
        # Save back to Excel while preserving sheet structure
        with pd.ExcelWriter(
            EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace"
        ) as writer:
          edited_df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

        st.cache_data.clear()
        st.success(
            "Inventory successfully updated and saved to Active supply"
            " list.xlsx!"
        )
        st.rerun()
      except Exception as e:
        st.error(f"Error saving to Excel file: {e}")


# ==========================================
# 3. ORDER SHEET (BELOW PAR)
# ==========================================
elif app_mode == "🛒 Order Sheet (Below Par)":
  st.subheader("On-Demand Reorder Sheet")
  st.markdown(
      "Items listed below have fallen **below their set PAR level** (Current"
      " Stock < PAR Level). The required order quantity is automatically"
      " calculated to restore stock back to PAR."
  )

  # Filter items where Current Stock < PAR Level
  order_df = df_data.copy()
  order_df["Current Stock"] = pd.to_numeric(
      order_df["Current Stock"], errors="coerce"
  ).fillna(0)
  order_df["PAR Level"] = pd.to_numeric(
      order_df["PAR Level"], errors="coerce"
  ).fillna(0)

  # Calculate needed quantity: PAR Level - Current Stock
  order_df["Quantity to Order"] = order_df["PAR Level"] - order_df["Current Stock"]
  below_par_df = order_df[order_df["Current Stock"] < order_df["PAR Level"]].copy()

  if below_par_df.empty:
    st.success(
        "🎉 All stock levels are currently at or above PAR! No items need to be"
        " ordered."
    )
  else:
    # Select and reorder columns highlighting Coupa UIN and Quantity to Order (Requirement 3)
    output_order_sheet = below_par_df[[
        "Item Description",
        "Coupa UIN",
        "PAR Level",
        "Current Stock",
        "Quantity to Order",
    ]]

    st.warning(
        f"Attention: {len(output_order_sheet)} item(s) require reordering."
    )
    st.dataframe(output_order_sheet, use_container_width=True, height=450)

    # Export option for ordering
    csv_data = output_order_sheet.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Download Order Sheet as CSV",
        data=csv_data,
        file_name="IPA1_Reorder_Sheet.csv",
        mime="text/csv",
    )


# ==========================================
# 4. ADD NEW ITEM TO INVENTORY
# ==========================================
elif app_mode == "➕ Add New Item":
  st.subheader("Add New Supply Item")
  st.markdown(
      "Add a new medical or safety supply item to the master database and Excel"
      " sheet."
  )

  with st.form("add_item_form"):
    new_desc = st.text_input("Item Description")
    new_uin = st.text_input("Coupa UIN")
    new_par = st.number_input("PAR Level", min_value=0, value=5, step=1)
    new_stock = st.number_input("Current Stock", min_value=0, value=0, step=1)

    add_submitted = st.form_submit_button("Add Item to Inventory")

    if add_submitted:
      if not new_desc or not new_uin:
        st.error("Please provide both Item Description and Coupa UIN.")
      else:
        new_row = {
            "Item Description": new_desc,
            "Coupa UIN": new_uin,
            "PAR Level": new_par,
            "Current Stock": new_stock,
            "Variance ": new_stock - new_par,
        }
        updated_df = pd.concat(
            [df_data, pd.DataFrame([new_row])], ignore_index=True
        )

        try:
          with pd.ExcelWriter(
              EXCEL_FILE,
              engine="openpyxl",
              mode="a",
              if_sheet_exists="replace",
          ) as writer:
            updated_df.to_excel(writer, sheet_name=SHEET_NAME, index=False)

          st.cache_data.clear()
          st.success(f"Successfully added '{new_desc}' to inventory!")
          st.rerun()
        except Exception as e:
          st.error(f"Error updating Excel: {e}")