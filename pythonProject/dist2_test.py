import pandas as pd
from openpyxl.styles import PatternFill
import streamlit as st
import io


@st.cache_data
def load_data(file_path):
    return pd.read_excel(file_path)

def check_duplicate_neptun_codes(data):
    duplicated = data[data.duplicated(subset=['Neptun kód'], keep=False)]
    if not duplicated.empty:
        print("Warning: There are students with duplicate Neptun kód:")
        for neptun_code in duplicated['Neptun kód'].unique():
            print(f"Neptun kód: {neptun_code}")
    else:
        print("No duplicate Neptun kód found.")

def highlight_exceeded_semesters(main_data, small_groups_data, semester_limits):
    # Merge the semester limits with the main data
    main_data = pd.merge(main_data, semester_limits, how='left', left_on='KépzésKód', right_on='Képzéskód')
    small_groups_data = pd.merge(small_groups_data, semester_limits, how='left', left_on='KépzésKód', right_on='Képzéskód')

    # Add additional semesters for "alapképzés" and "mesterképzés"
    main_data['Adjusted Félévszám'] = main_data['Félévszám'] + main_data['Képzési szint'].map({'alapképzés': 1, 'mesterképzés': 2}).fillna(0)
    small_groups_data['Adjusted Félévszám'] = small_groups_data['Félévszám'] + small_groups_data['Képzési szint'].map({'alapképzés': 1, 'mesterképzés': 2}).fillna(0)

    # Identify students exceeding the semester limits
    main_data['Exceed Limit'] = main_data['Aktív félévek'] > main_data['Adjusted Félévszám']
    small_groups_data['Exceed Limit'] = small_groups_data['Aktív félévek'] > small_groups_data['Adjusted Félévszám']

    return main_data, small_groups_data


def group_students(data):
    bins = [0, 2, 4, 6, 8, 10, 12]
    labels = ['1. éves', '2. éves', '3. éves', '4. éves', '5. éves', '6. éves']
    data['Évfolyam'] = pd.cut(data['Aktív félévek'], bins=bins, labels=labels, right=True)

    grouping_columns = ['KépzésNév', 'Képzési szint', 'Nyelv ID', 'Évfolyam']
    grouped = data.groupby(grouping_columns, observed=True).size().reset_index(name='Létszám')
    return grouped, data

def group_students_by_year(data):
    modified_data = data.copy()
    grouping_columns = ['KépzésNév', 'Képzési szint', 'Nyelv ID']

    unique_programs = modified_data[grouping_columns].drop_duplicates()

    for _, program in unique_programs.iterrows():
        program_mask = (
            (modified_data['KépzésNév'] == program['KépzésNév']) &
            (modified_data['Képzési szint'] == program['Képzési szint']) &
            (modified_data['Nyelv ID'] == program['Nyelv ID'])
        )
        program_data = modified_data[program_mask].copy()

        year_counts = program_data['Évfolyam'].value_counts().sort_index()
        years = list(year_counts.index)
        counts = list(year_counts.values)

        year_df = pd.DataFrame({'Évfolyam': years, 'Létszám': counts})
        year_df.sort_values('Évfolyam', inplace=True)
        year_df.reset_index(drop=True, inplace=True)

        year_order = {label: idx for idx, label in enumerate(year_df['Évfolyam'])}

        year_to_group = {year: year for year in year_df['Évfolyam']}

        merged_years = set()

        idx = 0
        while idx < len(year_df):
            current_year = year_df.loc[idx, 'Évfolyam']
            current_count = year_df.loc[idx, 'Létszám']

            if current_count >= 10:
                idx += 1
                continue

            prev_idx = idx - 1 if idx > 0 else None
            next_idx = idx + 1 if idx + 1 < len(year_df) else None

            candidates = []
            if prev_idx is not None:
                prev_year = year_df.loc[prev_idx, 'Évfolyam']
                prev_count = year_df.loc[prev_idx, 'Létszám']
                candidates.append({'idx': prev_idx, 'year': prev_year, 'count': prev_count})
            if next_idx is not None:
                next_year = year_df.loc[next_idx, 'Évfolyam']
                next_count = year_df.loc[next_idx, 'Létszám']
                candidates.append({'idx': next_idx, 'year': next_year, 'count': next_count})

            if not candidates:
                idx += 1
                continue

            min_count = min(c['count'] for c in candidates)
            min_candidates = [c for c in candidates if c['count'] == min_count]

            if len(min_candidates) > 1:
                merge_candidate = max(min_candidates, key=lambda c: year_order[c['year']])
            else:
                merge_candidate = min_candidates[0]

            merge_idx = merge_candidate['idx']
            merge_year = merge_candidate['year']

            year_df.loc[merge_idx, 'Létszám'] += current_count
            year_df.loc[idx, 'Létszám'] = 0

            target_group = year_to_group[merge_year]
            year_to_group[current_year] = target_group

            modified_data.loc[
                program_mask & (modified_data['Évfolyam'] == current_year),
                'Évfolyam'
            ] = target_group

            merged_years.add(current_year)
            year_df = year_df[year_df['Létszám'] > 0].reset_index(drop=True)
            idx = 0

        for merged_year in merged_years:
            target_group = year_to_group[merged_year]
            modified_data.loc[
                program_mask & (modified_data['Évfolyam'] == merged_year),
                'Évfolyam'
            ] = target_group

    return modified_data

def calculate_scholarship_index(data):
    data['Kredit szám'] = data['ElőzőFélévTeljesítettKredit'].apply(lambda x: min(x, 42))
    data['Ösztöndíjindex'] = data['Ösztöndíj átlag előző félév'] + ((data['Kredit szám'] / 27) - 1) / 2
    return data

def recalculate_year_for_small_groups(data):
    bins = [0, 2, 4, 6, 8, 10, 12]
    labels = ['1. éves', '2. éves', '3. éves', '4. éves', '5. éves', '6. éves']
    data['Évfolyam'] = pd.cut(data['Aktív félévek'], bins=bins, labels=labels, right=True)
    return data

def save_to_excel(main_data, separate_data):
    main_buffer = io.BytesIO()
    with pd.ExcelWriter(main_buffer, engine='openpyxl') as writer:
        main_data.to_excel(writer, index=False, sheet_name='MainData')
        apply_alternate_row_coloring(writer, main_data, 'MainData')
    main_buffer.seek(0)


    separate_buffer = io.BytesIO()
    with pd.ExcelWriter(separate_buffer, engine='openpyxl') as writer:
        separate_data.to_excel(writer, index=False, sheet_name='SeparateData')
        apply_alternate_row_coloring(writer, separate_data, 'SeparateData')
    separate_buffer.seek(0)

    return main_buffer, separate_buffer


def apply_alternate_row_coloring(writer, df, sheet_name):
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    grouping_columns = ['KépzésNév', 'Képzési szint', 'Nyelv ID', 'Évfolyam']
    df['TempGroupID'] = df[grouping_columns].apply(lambda x: ' | '.join(x.astype(str)), axis=1)
    last_row = df.shape[0] + 1
    last_col = df.shape[1]

    fill_colors = ['FFFFFF', 'D3D3D3']
    current_fill = 0
    previous_group = None
    red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

    for row in range(2, last_row + 1):
        group_id = df.iloc[row - 2]['TempGroupID']
        exceed_limit = df.iloc[row - 2]['Exceed Limit'] if 'Exceed Limit' in df.columns else False

        if exceed_limit:
            for col in range(1, last_col + 1):
                cell = worksheet.cell(row=row, column=col)
                cell.fill = red_fill
        else:
            if group_id != previous_group:
                current_fill = (current_fill + 1) % len(fill_colors)
                previous_group = group_id

            fill = PatternFill(start_color=fill_colors[current_fill], end_color=fill_colors[current_fill],
                               fill_type='solid')
            for col in range(1, last_col + 1):
                cell = worksheet.cell(row=row, column=col)
                cell.fill = fill

    df.drop(columns=['TempGroupID'], inplace=True)

def add_group_index(data):
    grouping_columns = ['KépzésNév', 'Képzési szint', 'Nyelv ID', 'Évfolyam']
    data['GroupID'] = data[grouping_columns].apply(lambda x: ' | '.join(x.astype(str)), axis=1)
    data['GroupIndex'] = (data['GroupID'] != data['GroupID'].shift()).cumsum()
    data.drop(columns=['GroupID'], inplace=True)
    cols = data.columns.tolist()
    cols.insert(0, cols.pop(cols.index('GroupIndex')))
    data = data[cols]
    return data

def filter_small_groups(data):
    grouping_columns = ['KépzésNév', 'Képzési szint', 'Nyelv ID']
    course_counts = data.groupby(grouping_columns).size().reset_index(name='Total_in_Course')
    data = data.merge(course_counts, on=grouping_columns)
    small_groups = data[data['Total_in_Course'] < 10]
    remaining_data = data[data['Total_in_Course'] >= 10]
    small_groups = small_groups.drop(columns=['Total_in_Course'])
    remaining_data = remaining_data.drop(columns=['Total_in_Course'])
    return remaining_data, small_groups

def sort_data(data):
    data = data.sort_values(by=['KépzésNév', 'Képzési szint', 'Nyelv ID', 'Évfolyam', 'Aktív félévek'], ascending=True)
    return data

def calculate_kodi(data):
    grouping_columns = ['KépzésNév', 'Képzési szint', 'Nyelv ID', 'Évfolyam']
    grouped = data.groupby(grouping_columns, observed=True)

    data['MinÖDI'] = grouped['Ösztöndíjindex'].transform('min')
    data['MaxÖDI'] = grouped['Ösztöndíjindex'].transform('max')

    def calculate_group_kodi(row):
        if row['MaxÖDI'] == row['MinÖDI']:
            return 100.0
        else:
            kodi = ((row['Ösztöndíjindex'] - row['MinÖDI']) / (row['MaxÖDI'] - row['MinÖDI'])) * 100
            return round(kodi, 6)

    data['KÖDI'] = data.apply(calculate_group_kodi, axis=1)

    max_odi_mask = data['Ösztöndíjindex'] == data['MaxÖDI']
    data.loc[max_odi_mask, 'KÖDI'] = 100.0

    min_odi_mask = data['Ösztöndíjindex'] == data['MinÖDI']
    data.loc[min_odi_mask, 'KÖDI'] = 0.0

    data.drop(columns=['MinÖDI', 'MaxÖDI'], inplace=True)

    return data

def remove_lower_kodi_duplicates(data):
    duplicated = data[data.duplicated(subset=['Neptun kód'], keep=False)]
    if not duplicated.empty:
        print("Removing lower KÖDI entries for duplicate Neptun kód:")
        for neptun_code in duplicated['Neptun kód'].unique():
            student_rows = data[data['Neptun kód'] == neptun_code]
            max_kodi_index = student_rows['KÖDI'].idxmax()
            indices_to_drop = student_rows.index.difference([max_kodi_index])
            data = data.drop(indices_to_drop)
            print(f"Neptun kód {neptun_code}: Kept index {max_kodi_index}, dropped indices {list(indices_to_drop)}")
    else:
        print("No duplicate Neptun kód found after KÖDI calculation.")
    return data.reset_index(drop=True)




def main():

    st.set_page_config(page_title="Step 1")

    st.title("Student Grouping")
    st.subheader("Upload an input file where 3,8 and 23 are filtered")
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    st.subheader("Upload max number of semesters file")
    uploaded_file2 = st.file_uploader("Choose an Excel file", type="xlsx")

    # Itt tartok, félév check jön

    if uploaded_file is not None and uploaded_file2 is not None:
        data = load_data(uploaded_file)
        semester_limits = load_data(uploaded_file2)
    else:
        st.stop()

    check_duplicate_neptun_codes(data)

    remaining_data, small_groups_data_initial = filter_small_groups(data)
    grouped_data, original_data = group_students(remaining_data)
    updated_data = group_students_by_year(remaining_data)
    updated_data = calculate_scholarship_index(updated_data)
    updated_data = calculate_kodi(updated_data)

    updated_data = remove_lower_kodi_duplicates(updated_data)

    # Re-run the grouping and redistribution after removing duplicates
    remaining_data, small_groups_data_after = filter_small_groups(updated_data)
    grouped_data, original_data = group_students(remaining_data)
    updated_data = group_students_by_year(remaining_data)
    updated_data = calculate_scholarship_index(updated_data)
    updated_data = calculate_kodi(updated_data)

    small_groups_data_combined = pd.concat([small_groups_data_initial, small_groups_data_after]).drop_duplicates().reset_index(drop=True)

    small_groups_data_combined = calculate_scholarship_index(small_groups_data_combined)

    updated_data = sort_data(updated_data)
    small_groups_data_combined = sort_data(small_groups_data_combined)

    updated_data = add_group_index(updated_data)
    small_groups_data_combined = add_group_index(small_groups_data_combined)

    updated_data, small_groups_data_combined = highlight_exceeded_semesters(updated_data, small_groups_data_combined,
                                                                            semester_limits)

    main_buffer, separate_buffer = save_to_excel(updated_data, small_groups_data_combined)

    st.subheader("Download Output Files")

    st.download_button(
        label="Download Main Data Excel",
        data=main_buffer.getvalue(),
        file_name="main_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.download_button(
        label="Download Small Groups Data Excel",
        data=separate_buffer.getvalue(),
        file_name="small_groups_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    main()