import streamlit as st
import pandas as pd
import numpy as np
from scipy.optimize import minimize_scalar
import matplotlib.pyplot as plt


@st.cache_data
def load_data(file_path):
    return pd.read_excel(file_path)


def get_user_inputs():
    st.sidebar.header("Adjustable Parameters")
    total_fund = st.sidebar.number_input("Total Scholarship Fund", value=100000, step=1000)
    group_percentage = st.sidebar.slider("Group Percentage (%)", min_value=1, max_value=100, value=30)
    max_amount_per_group = st.sidebar.number_input("Maximum Scholarship Amount per Student", value=2000, step=100)
    min_amount_per_group = st.sidebar.number_input("Minimum Scholarship Amount per Student", value=500, step=100)
    return total_fund, group_percentage / 100, max_amount_per_group, min_amount_per_group


def calculate_scholarship_amounts(data, gamma, max_amount_per_group, min_amount_per_group, group_percentage):
    scholarship_amounts = []
    groups = data['GroupID'].unique()

    for group in groups:
        group_data = data[data['GroupID'] == group].copy()
        total_students = len(group_data)
        num_recipients = int(np.ceil(group_percentage * total_students))

        # Sort students by KÖDI in descending order
        group_data = group_data.sort_values(by='KÖDI', ascending=False).reset_index(drop=True)

        # Select top students based on num_recipients
        recipients = group_data.iloc[:num_recipients].copy()

        # KÖDI cutoff (minimum KÖDI among recipients)
        KODI_cutoff = recipients['KÖDI'].min()

        # Handle case where KÖDI_cutoff == 100
        if KODI_cutoff == 100:
            KODI_cutoff = 99.999  # Slightly less than 100 to avoid division by zero

        # Normalize KÖDI scores to range between 0 and 1
        KODI_normalized = (recipients['KÖDI'] - KODI_cutoff) / (100 - KODI_cutoff)

        # Apply the exponential function
        f_K = KODI_normalized ** gamma

        # Calculate scholarship amounts
        recipients['Scholarship Amount'] = min_amount_per_group + f_K * (max_amount_per_group - min_amount_per_group)

        scholarship_amounts.append(recipients)

    # Combine all recipients
    all_recipients = pd.concat(scholarship_amounts, ignore_index=True)
    return all_recipients


def calculate_total_allocated_funds(recipients):
    total_allocated = recipients['Scholarship Amount'].sum()
    return total_allocated


def objective_function(gamma, data, max_amount_per_group, min_amount_per_group, group_percentage, total_fund):
    recipients = calculate_scholarship_amounts(data, gamma, max_amount_per_group, min_amount_per_group,
                                               group_percentage)
    total_allocated = calculate_total_allocated_funds(recipients)
    # Objective is the absolute difference between total_fund and total_allocated
    return abs(total_fund - total_allocated)


def optimize_gamma(data, max_amount_per_group, min_amount_per_group, group_percentage, total_fund):
    result = minimize_scalar(
        objective_function,
        bounds=(0.1, 5),
        args=(data, max_amount_per_group, min_amount_per_group, group_percentage, total_fund),
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
    required_columns = ['KépzésKód', 'KépzésNév', 'Neptun kód', 'Nyomtatási név', 'Képzési szint', 'Nyelv ID', 'Évfolyam', 'KÖDI']
    for col in required_columns:
        if col not in data.columns:
            st.error(f"Error: Column '{col}' not found in data.")
            return

    # Create a GroupID column to uniquely identify groups
    data['GroupID'] = data[['KépzésNév', 'Képzési szint', 'Nyelv ID', 'Évfolyam']].apply(
        lambda x: ' | '.join(x.astype(str)), axis=1)

    # Get user inputs
    total_fund, group_percentage, max_amount_per_group, min_amount_per_group = get_user_inputs()

    # Optimize gamma
    optimized_gamma = optimize_gamma(data, max_amount_per_group, min_amount_per_group, group_percentage, total_fund)

    # Recalculate scholarship amounts with optimized gamma
    recipients = calculate_scholarship_amounts(data, optimized_gamma, max_amount_per_group, min_amount_per_group,
                                               group_percentage)

    # Calculate total allocated funds
    total_allocated = calculate_total_allocated_funds(recipients)

    # Display results
    st.header("Results")
    st.write(f"**Optimized Gamma:** {optimized_gamma:.4f}")
    st.write(f"**Total Allocated Funds:** {total_allocated:.2f}")

    # Show recipients
    st.subheader("Scholarship Recipients")
    st.dataframe(recipients[['KépzésKód', 'KépzésNév', 'Neptun kód', 'Nyomtatási név', 'Képzési szint', 'Nyelv ID', 'Évfolyam', 'KÖDI']])

    # Visualize the distribution
    st.subheader("KÖDI vs. Scholarship Amount")
    visualize_distribution(recipients)


if __name__ == "__main__":
    main()
