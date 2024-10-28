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
    group_percentages = st.session_state.group_percentages

    with st.sidebar.expander("Set Group Percentages", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Add 5% to All"):
                for group in groups:
                    new_value = min(group_percentages[group] + 5, 100)
                    group_percentages[group] = new_value
                st.rerun()
        with col2:
            if st.button("Subtract 5% from All"):
                for group in groups:
                    new_value = max(group_percentages[group] - 5, 0)
                    group_percentages[group] = new_value
                st.rerun()

        for group in groups:
            percentage = st.number_input(
                f"Group {group} Percentage (%)",
                min_value=0,
                max_value=100,
                value=group_percentages[group],
                step=1,
                key=f"group_{group}"
            )
            group_percentages[group] = percentage

    group_percentages_decimal = {group: pct / 100 for group, pct in group_percentages.items()}
    return group_percentages_decimal


def calculate_scholarship_amounts_global(data, max_amount_per_group, min_amount_per_group, group_percentages, k, x0):
    global all_recipients
    recipients_list = []
    total_students = len(data)
    total_recipients = 0

    for group in data['GroupIndex'].unique():
        group_data = data[data['GroupIndex'] == group].copy()
        num_students_in_group = len(group_data)
        group_percentage = group_percentages.get(group, 0.3)
        num_recipients = int(np.ceil(group_percentage * num_students_in_group))

        group_data = group_data.sort_values(by='KÖDI', ascending=False).reset_index(drop=True)

        initial_recipients = group_data.iloc[:num_recipients].copy()

        if not initial_recipients.empty:
            last_included_KODI = initial_recipients['KÖDI'].iloc[-1]
            additional_recipients = group_data[group_data['KÖDI'] == last_included_KODI]
            all_recipients_group = pd.concat([initial_recipients, additional_recipients]).drop_duplicates()
            num_recipients_actual = len(all_recipients_group)
            total_recipients += num_recipients_actual

            recipients_list.append(all_recipients_group)
        else:
            continue

        all_recipients = pd.concat(recipients_list, ignore_index=True)
        all_recipients.drop_duplicates(inplace=True)

    KODI_cutoff_global = all_recipients['KÖDI'].min()

    epsilon = 0.01
    KODI_normalized = (all_recipients['KÖDI'] - KODI_cutoff_global) / (100 - KODI_cutoff_global + epsilon)
    KODI_normalized = np.clip(KODI_normalized, 0, 1)


    f_K = 1 / (1 + np.exp(-k * (KODI_normalized - x0)))

    all_recipients['Scholarship Amount'] = min_amount_per_group + f_K * (max_amount_per_group - min_amount_per_group)

    all_recipients['Scholarship Amount'] = all_recipients['Scholarship Amount'].round(2)

    cols = all_recipients.columns.tolist()
    cols.insert(0, cols.pop(cols.index('Scholarship Amount')))
    all_recipients = all_recipients[cols]
    return all_recipients, total_recipients, total_students

def calculate_total_allocated_funds(recipients):
    total_allocated = recipients['Scholarship Amount'].sum()
    return total_allocated

def objective_function_global(gamma, data, max_amount_per_group, min_amount_per_group, group_percentages, total_fund):
    recipients, _, _ = calculate_scholarship_amounts_global(
        data, gamma, max_amount_per_group, min_amount_per_group, group_percentages
    )
    total_allocated = calculate_total_allocated_funds(recipients)
    return abs(total_fund - total_allocated)

def visualize_distribution(recipients):
    plt.figure(figsize=(10, 6))
    plt.scatter(recipients['KÖDI'], recipients['Scholarship Amount'], alpha=0.7)
    plt.xlabel('KÖDI')
    plt.ylabel('Scholarship Amount')
    plt.title('KÖDI vs. Scholarship Amount')
    plt.grid(True)
    st.pyplot(plt)

def export_data_to_excel(recipients, required_columns):
    export_columns = required_columns + ['Scholarship Amount']

    recipients = recipients[export_columns]

    from io import BytesIO
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    recipients.to_excel(writer, index=False, sheet_name='Scholarship Recipients')
    writer.close()
    processed_data = output.getvalue()

    st.download_button(label='Download Excel File', data=processed_data,
                       file_name='Scholarship_Recipients.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def format_number_with_spaces(n):
    s = f"{n:,.2f}"
    s = s.replace(',', ' ')
    return s


def main():
    st.title("Scholarship Distribution Calculator")

    input_file = '/Users/bercelkovalik/Documents./InputOutput/output_data_test.xlsx'
    data = load_data(input_file)

    display_columns = ['KépzésNév','Neptun kód','Ösztöndíj átlag előző félév',
                                           'KÖDI', 'Scholarship Amount',]

    required_columns = ['GroupIndex', 'KépzésKód', 'KépzésNév', 'Neptun kód', 'Nyomtatási név',
                        'Felvétel féléve', 'Aktív félévek', 'Státusz2 jelen félév',
                        'Ösztöndíj átlag előző félév', 'Képzési szint', 'Nyelv ID', 'Tagozat',
                        'ElőzőFélévTeljesítettKredit', 'Hallgató kérvény azonosító', 'Évfolyam',
                        'Kredit szám', 'Ösztöndíjindex', 'KÖDI']

    export_columns = required_columns + ['Scholarship Amount']

    for col in required_columns:
        if col not in data.columns:
            st.error(f"Error: Column '{col}' not found in data.")
            return

    submitted_data = data[
        data['Hallgató kérvény azonosító'].notnull() & (data['Hallgató kérvény azonosító'] != '')].copy()

    submitted_data['GroupIndex'] = submitted_data['GroupIndex'].astype(int)
    data['GroupIndex'] = data['GroupIndex'].astype(int)

    groups = sorted(submitted_data['GroupIndex'].unique())

    if 'group_percentages' not in st.session_state:
        st.session_state.group_percentages = {group: 30 for group in groups}

    total_fund = st.sidebar.number_input("Total Scholarship Fund", value=100000000, step=1000)
    formatted_total_fund = format_number_with_spaces(total_fund)
    st.sidebar.write(f"Formatted Total Fund: {formatted_total_fund}")
    max_amount_per_group = st.sidebar.number_input("Maximum Scholarship Amount per Student", value=100000, step=100)
    min_amount_per_group = st.sidebar.number_input("Minimum Scholarship Amount per Student", value=30000, step=100)

    group_percentages = get_group_percentages(groups)

    total_students = len(data)
    total_recipients_estimated = 0
    for group in groups:
        num_students_in_group = len(submitted_data[submitted_data['GroupIndex'] == group])
        group_percentage = group_percentages.get(group, 0.3)
        num_recipients = int(np.ceil(group_percentage * num_students_in_group))
        total_recipients_estimated += num_recipients
    total_percentage_students = (total_recipients_estimated / total_students) * 100

    st.subheader("Adjust Logistic Function Parameters")
    k = st.number_input("Parameter k (steepness of the curve)", min_value=0.1, max_value=50.0, value=10.0, step=0.1)
    x0 = st.number_input("Parameter x₀ (midpoint of the curve)", min_value=0.0, max_value=1.0, value=0.5, step=0.01)

    recipients, total_recipients, total_students = calculate_scholarship_amounts_global(
        submitted_data, max_amount_per_group, min_amount_per_group, group_percentages, k, x0)

    total_allocated = calculate_total_allocated_funds(recipients)

    st.header("Results")
    if st.button("Export All Groups to Excel"):
        export_data_to_excel(recipients, required_columns)


    formatted_total_allocated = format_number_with_spaces(total_allocated)
    difference = total_allocated - total_fund
    formatted_difference = format_number_with_spaces(abs(difference))

    if total_allocated <= total_fund:
        st.markdown(
            f"<span style='color: green;'>**Total Allocated Funds:** {formatted_total_allocated} (Under by {formatted_difference})</span>",
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            f"<span style='color: red;'>**Total Allocated Funds:** {formatted_total_allocated} (Over by {formatted_difference})</span>",
            unsafe_allow_html=True
        )

    st.write(f"**Total Percentage of Students Receiving Scholarships:** {total_percentage_students:.2f}%")

    st.subheader("KÖDI vs. Scholarship Amount")
    visualize_distribution(recipients)

    st.subheader("Scholarship Recipients by Group")
    for group in groups:
        num_students_in_group = len(data[data['GroupIndex'] == group])

        group_recipients = recipients[recipients['GroupIndex'] == group]
        num_recipients_in_group = len(group_recipients)
        if not group_recipients.empty:
            st.markdown(
                f"### Group {group} (Total Students: {num_students_in_group}, Recipients: {num_recipients_in_group})")
            st.dataframe(group_recipients[display_columns])


if __name__ == "__main__":
    main()


