import pandas as pd


def load_data(file_path):
    return pd.read_excel(file_path)


def group_students(data):
    bins = [0, 2, 4, 6, 8, 10, 12]
    labels = ['1. éves', '2. éves', '3. éves', '4. éves', '5. éves', '6. éves']
    data['Évfolyam'] = pd.cut(data['Aktív félévek'], bins=bins, labels=labels, right=True)

    grouped = data.groupby(['KépzésKód', 'Évfolyam']).size().reset_index(name='Létszám')
    return grouped, data


def find_nearest_group(group, i):
    prev_idx, next_idx = i - 1, i + 1
    closest_idx = None
    while prev_idx >= 0 or next_idx < len(group):
        if prev_idx >= 0 and group.iloc[prev_idx]['Létszám'] >= 10:
            closest_idx = prev_idx if closest_idx is None else (
                prev_idx if group.iloc[prev_idx]['Létszám'] < group.iloc[closest_idx]['Létszám'] else closest_idx)
        if next_idx < len(group) and group.iloc[next_idx]['Létszám'] >= 10:
            closest_idx = next_idx if closest_idx is None else (
                next_idx if group.iloc[next_idx]['Létszám'] < group.iloc[closest_idx]['Létszám'] else closest_idx)
        if closest_idx is not None:
            return closest_idx
        prev_idx -= 1
        next_idx += 1
    return None


def redistribute_students(grouped, original_data):
    modified_data = original_data.copy()

    for code in grouped['KépzésKód'].unique():
        group = grouped[grouped['KépzésKód'] == code].copy()

        total_students = group['Létszám'].sum()
        if total_students < 10:
            continue

        for i in range(len(group)):
            if group.iloc[i]['Létszám'] < 10:
                current_year = group.iloc[i]['Évfolyam']
                original_index = modified_data[
                    (modified_data['KépzésKód'] == code) & (modified_data['Évfolyam'] == current_year)].index

                closest_year_index = find_nearest_group(group, i)

                if closest_year_index is not None:
                    new_year = group.iloc[closest_year_index]['Évfolyam']
                    modified_data.loc[original_index, 'Évfolyam'] = new_year
                    group.iloc[closest_year_index, group.columns.get_loc('Létszám')] += group.iloc[i]['Létszám']
                    group.iloc[i, group.columns.get_loc('Létszám')] = 0

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
    main_data.to_excel(output_file, index=False)
    separate_data.to_excel(separate_file, index=False)


def filter_small_groups(data):
    course_counts = data.groupby('KépzésKód').size().reset_index(name='Total_in_Course')
    data = data.merge(course_counts, on='KépzésKód')
    small_groups = data[data['Total_in_Course'] < 10]
    remaining_data = data[data['Total_in_Course'] >= 10]
    small_groups = small_groups.drop(columns=['Total_in_Course'])
    remaining_data = remaining_data.drop(columns=['Total_in_Course'])
    return remaining_data, small_groups


def sort_data(data):
    data = data.sort_values(by=['KépzésKód', 'Aktív félévek'], ascending=[True, True])
    return data


def calculate_kodi(data):
    grouped = data.groupby(['KépzésKód', 'Évfolyam'])
    data['MinÖDI'] = grouped['Ösztöndíjindex'].transform('min')
    data['MaxÖDI'] = grouped['Ösztöndíjindex'].transform('max')
    data['KÖDI'] = ((data['Ösztöndíjindex'] - data['MinÖDI']) / (data['MaxÖDI'] - data['MinÖDI'])) * 100
    data['KÖDI'] = data['KÖDI'].fillna(100)
    data.drop(columns=['MinÖDI', 'MaxÖDI'], inplace=True)

    return data


# Input és Output fájlok elnevezése
input_file = '/Users/bercelkovalik/Documents./InputOutput/Adatok.xlsx'
output_file = '/Users/bercelkovalik/Documents./InputOutput/output_data.xlsx'
separate_file = '/Users/bercelkovalik/Documents./InputOutput/small_groups_output.xlsx'

# Load
data = load_data(input_file)

# Függvények meghívása
remaining_data, small_groups_data = filter_small_groups(data)
grouped_data, original_data = group_students(remaining_data)
updated_data = redistribute_students(grouped_data, original_data)
updated_data_with_scholarship = calculate_scholarship_index(updated_data)
updated_data_with_scholarship_and_kodi = calculate_kodi(updated_data_with_scholarship)

# Rendezés
updated_data_with_scholarship = sort_data(updated_data_with_scholarship)
small_groups_data = recalculate_year_for_small_groups(small_groups_data)
small_groups_data_with_scholarship = calculate_scholarship_index(small_groups_data)
small_groups_data_with_scholarship = sort_data(small_groups_data_with_scholarship)

# Mentés
save_to_excel(updated_data_with_scholarship_and_kodi, small_groups_data, output_file, separate_file)

print("Process completed.")
