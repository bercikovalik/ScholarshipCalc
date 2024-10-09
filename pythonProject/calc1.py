import streamlit as st
import pandas as pd
import numpy as np
from scipy.optimize import minimize_scalar
import matplotlib.pyplot as plt

@st.cache_data
def load_data(file_path):
    data = pd.read_excel(file_path)
    return data

def get_group_percentages(groups):
    st.sidebar.header("Group Percentages")
    group_percentages = {}
    total_percentage = 0
    with st.sidebar.expander("Set Group Percentages"):
        for group in groups:
            # Assign a default percentage, e.g., 30%
            default_percentage = 35
            percentage = st.number_input(f"Group {group} Percentage (%)", min_value=0, max_value=100,
                                         value=default_percentage, step=1, key=f"group_{group}")
            group_percentages[group] = percentage / 100  # Convert to decimal
            total_percentage += percentage
    return group_percentages

def calculate_scholarship_amounts_global(data, gamma, max_amount_per_group, min_amount_per_group, group_percentages):
    recipients_list = []
    total_students = len(data)
    total_recipients = 0

    for group in data['GroupIndex'].unique():
        group_data = data[data['GroupIndex'] == group].copy()
        num_students_in_group = len(group_data)
        group_percentage = group_percentages.get(group, 0.3)  # Default to 30% if not specified
        num_recipients = int(np.ceil(group_percentage * num_students_in_group))

        # Sort students by KÖDI in descending order
        group_data = group_data.sort_values(by='KÖDI', ascending=False).reset_index(drop=True)

        # Select top students
        recipients = group_data.iloc[:num_recipients].copy()
        total_recipients += num_recipients

        recipients_list.append(recipients)

    # Combine all recipients
    all_recipients = pd.concat(recipients_list, ignore_index=True)

    # Global KÖDI Cutoff (minimum KÖDI among all recipients)
    KODI_cutoff_global = all_recipients['KÖDI'].min()

    # Handle case where KODI_cutoff_global == 100
    if KODI_cutoff_global == 100:
        KODI_cutoff_global = 99.999

    epsilon = 0.01  # Small positive value
    KODI_normalized = (all_recipients['KÖDI'] - KODI_cutoff_global) / (100 - KODI_cutoff_global + epsilon)

    # Ensure KODI_normalized is between 0 and 1
    KODI_normalized = np.clip(KODI_normalized, 0, 1)

    # Parameters for logistic function
    k = 10  # Adjust as needed
    x0 = 0.5  # Midpoint

    # Apply logistic function
    f_K = 1 / (1 + np.exp(-k * (KODI_normalized - x0)))

    # Calculate scholarship amounts
    all_recipients['Scholarship Amount'] = min_amount_per_group + f_K * (max_amount_per_group - min_amount_per_group)

    # Round scholarship amounts
    all_recipients['Scholarship Amount'] = all_recipients['Scholarship Amount'].round(2)

    # Reorder columns to place 'Scholarship Amount' first
    cols = all_recipients.columns.tolist()
    cols.insert(0, cols.pop(cols.index('Scholarship Amount')))
    all_recipients = all_recipients[cols]

    return all_recipients, total_recipients, total_students

def calculate_total_allocated_funds(recipients):
    total_allocated = recipients['Scholarship Amount'].sum()
    return total_allocated

def objective_function_global(gamma, data, max_amount_per_group, min_amount_per_group, group_percentages, total_fund):
    recipients, _, _ = calculate_scholarship_amounts_global(data, gamma, max_amount_per_group, min_amount_per_group, group_percentages)
    total_allocated = calculate_total_allocated_funds(recipients)
    return abs(total_fund - total_allocated)

def optimize_gamma_global(data, max_amount_per_group, min_amount_per_group, group_percentages, total_fund):
    result = minimize_scalar(
        objective_function_global,
        bounds=(0.1, 2),
        args=(data, max_amount_per_group, min_amount_per_group, group_percentages, total_fund),
        method='bounded'
    )
    optimized_gamma = result.x
    return optimized_gamma

def visualize_distribution(recipients):
    plt.figure(figsize=(10, 6))
    plt.scatter(recipients['KÖDI'], recipients['Scholarship Amount'], alpha=0.7)
    plt.xlabel('KÖDI')
    plt.ylabel('Scholarship Amount')
    plt.title('KÖDI vs. Scholarship Amount')
    plt.grid(True)
    st.pyplot(plt)


def main():
    st.title("Scholarship Distribution Calculator")

    # Load data
    input_file = '/Users/bercelkovalik/Documents./InputOutput/output_data.xlsx'
    data = load_data(input_file)

    # Ensure that the necessary columns are present
    required_columns = ['GroupIndex', 'KépzésKód', 'KépzésNév', 'Neptun kód', 'Nyomtatási név',
                        'Képzési szint', 'Nyelv ID', 'Évfolyam', 'KÖDI']
    for col in required_columns:
        if col not in data.columns:
            st.error(f"Error: Column '{col}' not found in data.")
            return

    # Convert GroupIndex to integer if necessary
    data['GroupIndex'] = data['GroupIndex'].astype(int)

    # Get list of groups
    groups = sorted(data['GroupIndex'].unique())

    # Get user inputs
    total_fund = st.sidebar.number_input("Total Scholarship Fund", value=100000000, step=1000)
    max_amount_per_group = st.sidebar.number_input("Maximum Scholarship Amount per Student", value=100000, step=100)
    min_amount_per_group = st.sidebar.number_input("Minimum Scholarship Amount per Student", value=30000, step=100)

    # Get group percentages
    group_percentages = get_group_percentages(groups)

    # Calculate total percentage of students receiving scholarships
    total_students = len(data)
    total_recipients_estimated = 0
    for group in groups:
        num_students_in_group = len(data[data['GroupIndex'] == group])
        group_percentage = group_percentages.get(group, 0.3)  # Default to 30% if not specified
        num_recipients = int(np.ceil(group_percentage * num_students_in_group))
        total_recipients_estimated += num_recipients
    total_percentage_students = (total_recipients_estimated / total_students) * 100

    # Optimize gamma
    optimized_gamma = optimize_gamma_global(data, max_amount_per_group, min_amount_per_group, group_percentages,
                                            total_fund)

    # Calculate scholarship amounts using global functions
    recipients, total_recipients, total_students = calculate_scholarship_amounts_global(
        data, optimized_gamma, max_amount_per_group, min_amount_per_group, group_percentages)

    # Calculate total allocated funds
    total_allocated = calculate_total_allocated_funds(recipients)

    # Display results
    st.header("Results")
    st.write(f"**Optimized Gamma:** {optimized_gamma:.4f}")
    difference = total_allocated - total_fund
    if difference <= 0:
        st.markdown(
            f"<span style='color: green;'>**Total Allocated Funds:** {total_allocated:.2f} (Under by {abs(difference):.2f})</span>",
            unsafe_allow_html=True)
    else:
        st.markdown(
            f"<span style='color: red;'>**Total Allocated Funds:** {total_allocated:.2f} (Over by {difference:.2f})</span>",
            unsafe_allow_html=True)

    st.write(f"**Total Percentage of Students Receiving Scholarships:** {total_percentage_students:.2f}%")

    # Visualize the distribution
    st.subheader("KÖDI vs. Scholarship Amount")
    visualize_distribution(recipients)

    # Display recipients grouped by GroupIndex
    st.subheader("Scholarship Recipients by Group")
    for group in groups:
        # Get the original number of students in the group
        num_students_in_group = len(data[data['GroupIndex'] == group])

        group_recipients = recipients[recipients['GroupIndex'] == group]
        num_recipients_in_group = len(group_recipients)
        if not group_recipients.empty:
            st.markdown(
                f"### Group {group} (Total Students: {num_students_in_group}, Recipients: {num_recipients_in_group})")
            st.dataframe(group_recipients[['GroupIndex', 'Scholarship Amount', 'KépzésKód', 'KépzésNév',
                                           'Neptun kód', 'Nyomtatási név', 'Képzési szint',
                                           'Nyelv ID', 'Évfolyam', 'KÖDI']])




if __name__ == "__main__":
    main()


