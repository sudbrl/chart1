import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st

# Streamlit app
st.title("Branch Loan Report Generator")

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
    bars_branch = ax_branch.bar(combined_branches['BranchName'], combined_branches['Change'], 
                                 color=colors_branch, edgecolor='black', alpha=0.7)
    ax_branch.set_title('Top 5 Increasing and Declining Branches (in Crores)', fontsize=16)
    ax_branch.set_xlabel('Branch Name', fontsize=14)
    ax_branch.set_ylabel('Loan Change (in currency)', fontsize=14)
    ax_branch.set_xticklabels(combined_branches['BranchName'], rotation=45, ha='right')
    
    for bar, change_crore in zip(bars_branch, combined_branches['Change (Crore)']):
        yval = bar.get_height()
        ax_branch.text(bar.get_x() + bar.get_width() / 2, yval, f'Rs.{abs(change_crore):.2f} Cr', 
                       ha='center', va='bottom', fontsize=10)

    st.subheader("Branch Loan Change Analysis")
    st.pyplot(fig_branch)

    # Display the data table
    st.dataframe(combined_branches)
