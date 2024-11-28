import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
###Copyright 2024, Bercel Kovalik-Deák, All rights reserved

@st.cache_data
def load_data(file_path):
    data = pd.read_excel(file_path)
    return data


def get_group_percentages(groups):
    st.sidebar.header("Group Percentages")

    # Initialize group_percentages in session state if it doesn't exist or if it's missing groups
    if 'group_percentages' not in st.session_state:
        st.session_state.group_percentages = {group: 30 for group in groups}
    else:
        # Ensure every group in `groups` has an entry in `group_percentages`
        for group in groups:
            if group not in st.session_state.group_percentages:
                st.session_state.group_percentages[group] = 30

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

        # Add number input for each group
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

def calculate_scholarship_amounts_global(submitted_data, all_data, max_amount_per_group, min_amount_per_group, group_percentages, k, x0):
    global all_recipients
    recipients_list = []
    total_students = len(all_data)
    total_recipients = 0
    group_min_kodi_dict = {}
    group_min_index_dict = {}

    for group in all_data['GroupIndex'].unique():
        all_group_data = all_data[all_data['GroupIndex'] == group].copy()
        group_submitted_data = submitted_data[submitted_data['GroupIndex'] == group].copy()

        num_students_in_group = len(all_group_data)
        group_percentage = group_percentages.get(group, 0.3)

        num_recipients = int(np.ceil(group_percentage * num_students_in_group))

        group_submitted_data = group_submitted_data.sort_values(by='KÖDI', ascending=False).reset_index(drop=True)

        initial_recipients = group_submitted_data.iloc[:num_recipients].copy()

        if not initial_recipients.empty:
            last_included_KODI = initial_recipients['KÖDI'].iloc[-1]

            additional_recipients = group_submitted_data[
                (group_submitted_data['KÖDI'] == last_included_KODI) & (group_submitted_data.index >= num_recipients)]

            all_recipients_group = pd.concat([initial_recipients, additional_recipients]).drop_duplicates(
                subset=['Neptun kód'])

            if len(all_recipients_group) < num_recipients:
                remaining_students = group_submitted_data.loc[
                    ~group_submitted_data['Neptun kód'].isin(all_recipients_group['Neptun kód'])]
                num_needed = num_recipients - len(all_recipients_group)
                additional_needed = remaining_students.iloc[:num_needed]
                all_recipients_group = pd.concat([all_recipients_group, additional_needed])

            num_recipients_actual = len(all_recipients_group)
            total_recipients += num_recipients_actual

            group_min_kodi_dict[group] = last_included_KODI
            group_min_index_dict[group] = all_recipients_group['Ösztöndíjindex'].min()

            all_recipients_group['Group Minimum Ösztöndíjindex'] = all_recipients_group['Ösztöndíjindex'].min()
            recipients_list.append(all_recipients_group)

    all_recipients = pd.concat(recipients_list, ignore_index=True)
    all_recipients.drop_duplicates(inplace=True)

    KODI_cutoff_global = all_recipients['KÖDI'].min()

    epsilon = 0.01
    KODI_normalized = (all_recipients['KÖDI'] - KODI_cutoff_global) / (100 - KODI_cutoff_global + epsilon)
    KODI_normalized = np.clip(KODI_normalized, 0, 1)

    f_K = 1 / (1 + np.exp(-k * (KODI_normalized - x0)))

    all_recipients['Scholarship Amount'] = min_amount_per_group + f_K * (max_amount_per_group - min_amount_per_group)
    all_recipients.loc[all_recipients['KÖDI'] == 100, 'Scholarship Amount'] = max_amount_per_group
    all_recipients['Scholarship Amount'] = (all_recipients['Scholarship Amount'] / 100).round() * 100

    cols = all_recipients.columns.tolist()
    cols.insert(0, cols.pop(cols.index('Scholarship Amount')))
    cols.insert(1, cols.pop(cols.index('Group Minimum Ösztöndíjindex')))
    all_recipients = all_recipients[cols]

    return all_recipients, total_recipients, total_students, group_min_index_dict

def calculate_total_allocated_funds(recipients):
    total_allocated = recipients['Scholarship Amount'].sum()
    return total_allocated

def visualize_distribution(recipients):
    plt.figure(figsize=(10, 6))
    plt.scatter(recipients['KÖDI'], recipients['Scholarship Amount'], alpha=0.7)
    plt.xlabel('KÖDI')
    plt.ylabel('Scholarship Amount')
    plt.title('KÖDI vs. Scholarship Amount')
    plt.grid(True)
    st.pyplot(plt)

def export_data_to_excel(data_to_export, required_columns):
    export_columns = required_columns + ['Scholarship Amount', 'Group Minimum Ösztöndíjindex']

    data_to_export = data_to_export[export_columns]

    from io import BytesIO
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    data_to_export.to_excel(writer, index=False, sheet_name='Scholarship Recipients')
    writer.close()
    processed_data = output.getvalue()

    st.download_button(label='Download Excel File', data=processed_data,
                       file_name='Scholarship_Data.xlsx',
                       mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

def format_number_with_spaces(n):
    s = f"{n:,.2f}"
    s = s.replace(',', ' ')
    return s


def main():
    st.set_page_config(page_title="Step 2")
    st.title("Scholarship Distribution Calculator")

    st.subheader("Upload Input File")
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file is not None:
        data = load_data(uploaded_file)
    else:
        st.stop()

    display_columns = ['KépzésNév','Neptun kód','Ösztöndíj átlag előző félév',
                                           'KÖDI', 'Scholarship Amount']

    required_columns = ['GroupIndex', 'KépzésKód', 'KépzésNév', 'Neptun kód', 'Nyomtatási név',
                        'Felvétel féléve', 'Aktív félévek', 'Státusz2 jelen félév',
                        'Ösztöndíj átlag előző félév', 'Képzési szint_x', 'Nyelv ID', 'Tagozat_x',
                        'ElőzőFélévTeljesítettKredit', 'Hallgató kérvény azonosító', 'Évfolyam',
                        'Kredit szám', 'Ösztöndíjindex', 'KÖDI', 'Exceed Limit']

    for col in required_columns:
        if col not in data.columns:
            st.error(f"Error: Column '{col}' not found in data.")
            return
    ###DEBUG
    st.write(f"Debug: Data type of 'Exceed Limit' column: {data['Exceed Limit'].dtype}")
    st.write("Debug: Unique values in 'Exceed Limit' column:", data['Exceed Limit'].unique())
    ###DEBUG END

    submitted_data_all = data[data['Exceed Limit'] == False].copy()
    submitted_data_over = data[data['Exceed Limit'] == True].copy()
    submitted_data_kerveny = submitted_data_all[
        submitted_data_all['Hallgató kérvény azonosító'].notnull() & (submitted_data_all['Hallgató kérvény azonosító'] != '')].copy()


    submitted_data_all['GroupIndex'] = submitted_data_all['GroupIndex'].astype(int)
    submitted_data_kerveny['GroupIndex'] = submitted_data_kerveny['GroupIndex'].astype(int)

    groups = sorted(submitted_data_all['GroupIndex'].unique())

    if 'group_percentages' not in st.session_state:
        st.session_state.group_percentages = {group: 30 for group in groups}

    total_fund = st.sidebar.number_input("Total Scholarship Fund", value=100000000, step=1000)
    formatted_total_fund = format_number_with_spaces(total_fund)
    st.sidebar.write(f"Formatted Total Fund: {formatted_total_fund}")
    max_amount_per_group = st.sidebar.number_input("Maximum Scholarship Amount per Student", value=100000, step=100)
    min_amount_per_group = st.sidebar.number_input("Minimum Scholarship Amount per Student", value=30000, step=100)

    group_percentages = get_group_percentages(groups)

    total_students = len(submitted_data_all)
    total_recipients_estimated = 0
    for group in groups:
        num_students_in_group_all = len(submitted_data_all[submitted_data_all['GroupIndex'] == group])
        group_percentage = group_percentages.get(group, 0.3)
        num_recipients = int(np.ceil(group_percentage * num_students_in_group_all))
        total_recipients_estimated += num_recipients
    total_percentage_students = (total_recipients_estimated / total_students) * 100

    st.subheader("Adjust Logistic Function Parameters")
    k = st.number_input("Parameter k (steepness of the curve)", min_value=0.1, max_value=50.0, value=10.0, step=0.1)
    x0 = st.number_input("Parameter x₀ (midpoint of the curve)", min_value=0.0, max_value=1.0, value=0.5, step=0.01)

    recipients, total_recipients, total_students, group_min_kodi_dict= calculate_scholarship_amounts_global(
        submitted_data_kerveny, submitted_data_all, max_amount_per_group, min_amount_per_group, group_percentages, k, x0)

    total_allocated = calculate_total_allocated_funds(recipients)

    group_min_kodi_df = pd.DataFrame(list(group_min_kodi_dict.items()), columns=['GroupIndex', 'Group Minimum Ösztöndíjindex'])

    data = pd.merge(data, group_min_kodi_df, on='GroupIndex', how='left')

    all_students_data = pd.merge(
        data,
        recipients[['Neptun kód', 'Scholarship Amount']],
        on='Neptun kód',
        how='left'
    )
    all_students_data['Scholarship Amount'] = all_students_data['Scholarship Amount'].fillna('')
    all_students_data['Group Minimum Ösztöndíjindex'] = all_students_data['Group Minimum Ösztöndíjindex'].fillna('')

    all_students_data['Exceeded Semester Limit'] = all_students_data['Neptun kód'].isin(
        submitted_data_over['Neptun kód'])
    all_students_data['Scholarship Amount'] = all_students_data['Scholarship Amount'].fillna('Not Eligible')

    st.header("Results")
    if st.button("Export All Students to Excel"):
        export_data_to_excel(all_students_data, required_columns)


    formatted_total_allocated = format_number_with_spaces(total_allocated)
    difference = total_allocated - total_fund
    formatted_difference = format_number_with_spaces(abs(difference))


    header = st.container()

    header.write("""<div class='fixed-header'/>""", unsafe_allow_html=True)
    if total_allocated <= total_fund:
        header.markdown(
            f"<span style='color: green;'>**Total Allocated Funds:** {formatted_total_allocated} (Under by {formatted_difference})</span>",
            unsafe_allow_html=True
        )
    else:
        header.markdown(
            f"<span style='color: red;'>**Total Allocated Funds:** {formatted_total_allocated} (Over by {formatted_difference})</span>",
            unsafe_allow_html=True
        )
    ### Custom CSS for the sticky header
    st.markdown(
        """
    <style>
        div[data-testid="stVerticalBlock"] div:has(div.fixed-header) {
            position: sticky;
            top: 2.875rem;
            background-color: white;
            z-index: 999;
        }
        .fixed-header {
            border-bottom: 1px solid black;
        }
    </style>
        """,
        unsafe_allow_html=True
    )
    submitted_data_kerveny_num = len(submitted_data_kerveny)
    kerveny_percentage = (submitted_data_kerveny_num / total_students) * 100
    st.write(f"**Total Percentage of Students Receiving Scholarships:** {total_percentage_students:.2f}%")
    st.write(f"Total Number of students who submitted request: **{submitted_data_kerveny_num}** out of **{total_students}**. Percentage: **{kerveny_percentage:.2f}**%")

    st.subheader("KÖDI vs. Scholarship Amount")
    visualize_distribution(recipients)

    st.subheader("Scholarship Recipients by Group")
    for group in groups:
        num_students_in_group_all = len(data[data['GroupIndex'] == group])

        group_recipients = recipients[recipients['GroupIndex'] == group]
        num_recipients_in_group = len(group_recipients)
        actual_percentage = (num_recipients_in_group / num_students_in_group_all) * 100
        if not group_recipients.empty:
            st.markdown(
                f"### Group {group} (Total Students: {num_students_in_group_all}, Recipients: {num_recipients_in_group}, Actual percentage: {actual_percentage:.2f}%)")
            st.dataframe(group_recipients[display_columns])


if __name__ == "__main__":
    main()


