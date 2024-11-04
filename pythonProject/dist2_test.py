import pandas as pd
from openpyxl.styles import PatternFill

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

def save_to_excel(main_data, separate_data, output_file, separate_file):
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        main_data.to_excel(writer, index=False, sheet_name='MainData')
        apply_alternate_row_coloring(writer, main_data, 'MainData')

    with pd.ExcelWriter(separate_file, engine='openpyxl') as writer:
        separate_data.to_excel(writer, index=False, sheet_name='SeparateData')
        apply_alternate_row_coloring(writer, separate_data, 'SeparateData')

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

    for row in range(2, last_row + 1):
        group_id = df.iloc[row - 2]['TempGroupID']
        if group_id != previous_group:
            current_fill = (current_fill + 1) % len(fill_colors)
            previous_group = group_id

        fill = PatternFill(start_color=fill_colors[current_fill], end_color=fill_colors[current_fill], fill_type='solid')
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

# Input and Output file paths
input_file = '/Users/bercelkovalik/Documents./InputOutput/Adatok.xlsx'
output_file = '/Users/bercelkovalik/Documents./InputOutput/output_data_test.xlsx'
separate_file = '/Users/bercelkovalik/Documents./InputOutput/small_groups_output_test.xlsx'


data = load_data(input_file)

# Check for duplicate Neptun kód entries before processing
check_duplicate_neptun_codes(data)

# Initial Function Calls
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

# Combine small groups data
small_groups_data_combined = pd.concat([small_groups_data_initial, small_groups_data_after]).drop_duplicates().reset_index(drop=True)

# Sorting
updated_data = sort_data(updated_data)
small_groups_data_combined = sort_data(small_groups_data_combined)

# Add GroupIndex after sorting
updated_data = add_group_index(updated_data)
small_groups_data_combined = add_group_index(small_groups_data_combined)

# Save to Excel
save_to_excel(updated_data, small_groups_data_combined, output_file, separate_file)


print("Process completed.")
