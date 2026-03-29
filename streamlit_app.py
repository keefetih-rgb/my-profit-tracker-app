import streamlit as st
import pandas as pd
import io

# Sets the look of the website
st.set_page_config(page_title="Profit Calculator", layout="wide")

st.title("💰 E-commerce Profit Spreadsheet Generator")
st.write("Edit the table below, then click 'Download' to get an Excel file with working formulas!")

# 1. Create a starting table
if 'data' not in st.session_state:
    st.session_state.data = pd.DataFrame([
        {"Product": "Item A", "Price": 50.0, "Cost": 20.0, "Platform": "Amazon", "Comm_Pct": 0.15, "Fee": 0.0},
        {"Product": "Item B", "Price": 30.0, "Cost": 10.0, "Platform": "eBay", "Comm_Pct": 0.13, "Fee": 0.30},
    ])

# 2. Show the interactive table to the user
edited_df = st.data_editor(st.session_state.data, num_rows="dynamic")

# 3. The "Magic" Button
if st.button("Generate My Excel File"):
    output = io.BytesIO()
    
    # This part writes the actual Excel file
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        edited_df.to_excel(writer, index=False, sheet_name='Sheet1')
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']

        # Add some pretty formatting
        money_fmt = workbook.add_format({'num_format': '$#,##0.00'})
        pct_fmt = workbook.add_format({'num_format': '0%'})

        # Write the formulas into the hidden spreadsheet cells
        # We do this for 100 rows so you can add more items later in Excel
        for row in range(2, 101):
            # Comm $ = (Price * Pct) + Fee
            worksheet.write_formula(f'G{row}', f'=(B{row}*E{row})+F{row}', money_fmt)
            # Profit = Price - Cost - Comm $
            worksheet.write_formula(f'H{row}', f'=B{row}-C{row}-G{row}', money_fmt)
            # Margin % = Profit / Price
            worksheet.write_formula(f'I{row}', f'=IFERROR(H{row}/B{row}, 0)', pct_fmt)

    st.success("✅ Ready!")
    st.download_button(
        label="📥 Click here to Download",
        data=output.getvalue(),
        file_name="Ecommerce_Profit_Tracker.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
