import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO
from fpdf import FPDF
import streamlit as st
import tempfile

# Streamlit app
st.title("Comprehensive Loan Report Generator")

# File uploader
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    # Load the uploaded Excel file
    file_path = uploaded_file

    # Branch Sheet Analysis
    branch_sheet = 'Branch'
    branch_df = pd.read_excel(file_path, sheet_name=branch_sheet)

    # Sort data for top 5 increasing and declining branches
    top_5_increasing = branch_df.nlargest(5, 'Change')
    top_5_declining = branch_df.nsmallest(5, 'Change')

    # Combine the top increasing and declining branches
    combined_branches = pd.concat([top_5_increasing[['BranchName', 'Change']],
                                   top_5_declining[['BranchName', 'Change']]])
    combined_branches['Change (Crore)'] = combined_branches['Change'] / 1e7

    # Plot the bar chart for Branch data
    fig_branch, ax_branch = plt.subplots(figsize=(12, 8))
    colors_branch = ['green' if x > 0 else 'red' for x in combined_branches['Change']]
    bars_branch = ax_branch.bar(combined_branches['BranchName'], combined_branches['Change'], color=colors_branch, edgecolor='black', alpha=0.7)
    ax_branch.set_title('Top 5 Increasing and Declining Branches (in Crores)', fontsize=16)
    ax_branch.set_xlabel('Branch Name', fontsize=14)
    ax_branch.set_ylabel('Loan Change (in currency)', fontsize=14)
    ax_branch.set_xticklabels(combined_branches['BranchName'], rotation=45, ha='right')
    for bar, change_crore in zip(bars_branch, combined_branches['Change (Crore)']):
        yval = bar.get_height()
        ax_branch.text(bar.get_x() + bar.get_width() / 2, yval, f'Rs.{abs(change_crore):.2f} Cr', ha='center', va='bottom', fontsize=10)
    st.subheader("Branch Loan Change Analysis")
    st.pyplot(fig_branch)

    # Compare Sheet Analysis
    compare_sheet = 'Compare'
    df_compare = pd.read_excel(file_path, sheet_name=compare_sheet)
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
    st.dataframe(df_compare_sorted)

    # Generate PDF Report
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Comprehensive Loan Report", ln=True, align='C')

    # Save the Branch chart to a temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile:
        fig_branch.savefig(tmpfile.name, format="png")
        tmpfile.close()
        pdf.image(tmpfile.name, x=10, y=30, w=190)

    pdf.add_page()
    pdf.cell(200, 10, txt="Loan Type Balance Change", ln=True, align='C')
    with tempfile.NamedTemporaryFile(delete=False, suffix=".png") as tmpfile:
        fig_compare.savefig(tmpfile.name, format="png")
        tmpfile.close()
        pdf.image(tmpfile.name, x=10, y=30, w=190)

    # Save PDF to a buffer
    pdf_buffer = BytesIO()
    pdf_output = pdf.output(dest='S').encode('latin1')
    pdf_buffer.write(pdf_output)
    pdf_buffer.seek(0)

    st.download_button(label="Download Comprehensive PDF Report", data=pdf_buffer, file_name="Comprehensive_Loan_Report.pdf", mime="application/pdf")
