import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from fpdf import FPDF
import streamlit as st
import tempfile
import os

# Streamlit app
st.title("Comprehensive Loan Report Generator")

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        # Read uploaded file into BytesIO
        file_bytes = BytesIO(uploaded_file.read())

        # Branch Sheet Analysis
        branch_sheet = 'Branch'
        branch_df = pd.read_excel(file_bytes, sheet_name=branch_sheet)

        # Change the required columns to 'index' and 'Percent Change'
        required_branch_cols = {'index', 'Percent Change'}
        if not required_branch_cols.issubset(branch_df.columns):
            st.error(f"Branch sheet must contain columns: {required_branch_cols}")
        else:
            # Set 'index' as the index in the DataFrame, and use 'Percent Change' for sorting
            branch_df.set_index('index', inplace=True)
            top_5_increasing = branch_df.nlargest(5, 'Percent Change')
            top_5_declining = branch_df.nsmallest(5, 'Percent Change')

            # Combine the top increasing and declining branches
            combined_branches = pd.concat([
                top_5_increasing[['Percent Change']],
                top_5_declining[['Percent Change']]
            ])
            combined_branches['Change (Percent)'] = combined_branches['Percent Change']

            # Plot the bar chart for Branch data
            fig_branch, ax_branch = plt.subplots(figsize=(12, 8))
            colors_branch = ['green' if x > 0 else 'red' for x in combined_branches['Percent Change']]
            bars_branch = ax_branch.bar(combined_branches.index, combined_branches['Percent Change'],
                                        color=colors_branch, edgecolor='black', alpha=0.7)
            ax_branch.set_title('Top 5 Increasing and Declining Branches (in Percent)', fontsize=16)
            ax_branch.set_xlabel('Branch Name', fontsize=14)
            ax_branch.set_ylabel('Percent Change', fontsize=14)
            ax_branch.set_xticks(range(len(combined_branches.index)))
            ax_branch.set_xticklabels(combined_branches.index, rotation=45, ha='right')
            for bar, change_percent in zip(bars_branch, combined_branches['Change (Percent)']):
                yval = bar.get_height()
                ax_branch.text(bar.get_x() + bar.get_width() / 2, yval, f'{change_percent:.2f}%',
                               ha='center', va='bottom', fontsize=10)
            st.subheader("Branch Loan Change Analysis (Percent Change)")
            st.pyplot(fig_branch)

        # Move back to start of BytesIO
        file_bytes.seek(0)

        # Compare Sheet Analysis
        compare_sheet = 'Compare'
        df_compare = pd.read_excel(file_bytes, sheet_name=compare_sheet)

        required_compare_cols = {'index', 'Change'}
        if not required_compare_cols.issubset(df_compare.columns):
            st.error(f"Compare sheet must contain columns: {required_compare_cols}")
        else:
            df_compare['Change_Crore_NRs'] = df_compare['Change'] / 1e7
            df_compare_sorted = df_compare.sort_values('Change_Crore_NRs', ascending=False)

            colors_compare = [
                'green' if i < 3 else 'blue' if i < len(df_compare_sorted) - 3 else 'red'
                for i in range(len(df_compare_sorted))
            ]

            fig_compare, ax_compare = plt.subplots(figsize=(12, 8))
            ax_compare.barh(df_compare_sorted['index'], df_compare_sorted['Change_Crore_NRs'], color=colors_compare)
            ax_compare.set_xlabel('Change in Balance (Crore NRs)')
            ax_compare.set_ylabel('Loan Type')
            ax_compare.set_title('Change in Loan Balances Across Loan Types (in Crore NRs)')
            ax_compare.grid(axis='x', linestyle='--', alpha=0.7)
            st.subheader("Loan Type Balance Change Analysis")
            st.pyplot(fig_compare)

        # Generate PDF Report
        pdf = FPDF()
        pdf.set_auto_page_break(auto=True, margin=15)
        pdf.add_page()
        pdf.set_font("Arial", 'B', 16)
        pdf.cell(200, 10, txt="Comprehensive Loan Report", ln=True, align='C')
        pdf.ln(10)

        if 'fig_branch' in locals():
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile:
                fig_branch.savefig(tmpfile.name, format="png")
                pdf.image(tmpfile.name, x=10, y=30, w=190)
                os.unlink(tmpfile.name)

        pdf.add_page()
        pdf.cell(200, 10, txt="Loan Type Balance Change", ln=True, align='C')
        pdf.ln(10)

        if 'fig_compare' in locals():
            with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile:
                fig_compare.savefig(tmpfile.name, format="png")
                pdf.image(tmpfile.name, x=10, y=30, w=190)
                os.unlink(tmpfile.name)

        # Save PDF to buffer
        pdf_buffer = BytesIO()
        pdf_output = pdf.output(dest='S').encode('latin1')
        pdf_buffer.write(pdf_output)
        pdf_buffer.seek(0)

        st.download_button(label="Download Comprehensive PDF Report",
                           data=pdf_buffer,
                           file_name="Comprehensive_Loan_Report.pdf",
                           mime="application/pdf")

    except Exception as e:
        st.error(f"Error processing file: {e}")
