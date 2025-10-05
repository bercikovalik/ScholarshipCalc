import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl.styles import PatternFill
from pygments.unistring import combine


def calculate_summary(combined_df):
    """Calculates the average scholarship summary.

    Args:
        combined_df: The DataFrame from the original processing.

    Returns:
        A DataFrame representing the summary table, or None on error.
    """
    if 'Ösztöndíjindex' not in combined_df.columns or \
       'KépzésNév' not in combined_df.columns or \
       '1 havi Ösztöndíj' not in combined_df.columns:
        st.error("Error: Required columns ('Ösztöndíjindex', 'KépzésNév', '1 havi Ösztöndíj') are missing for summary calculation.")
        return None

    combined_df['CombinedKey'] = combined_df['KépzésNév'].astype(str) + ", " + \
                                 combined_df['Nyelv ID'].astype(str) + ", " + \
                                 combined_df['Képzési szint_x'].astype(str)

    combined_df['Rounded Ösztöndíjindex'] = combined_df['Ösztöndíjindex'].round(2)
    summary_df = combined_df.groupby(['Rounded Ösztöndíjindex', 'CombinedKey'])['1 havi Ösztöndíj'].mean().reset_index()

    pivot_df = summary_df.pivot_table(index='Rounded Ösztöndíjindex', columns='CombinedKey', values='1 havi Ösztöndíj',
                                      fill_value=0)
    pivot_df['Average of Courses'] = pivot_df.apply(lambda row: row[row != 0].mean(), axis=1)

    return pivot_df

def main():
    st.title("Final Step: Merge Scholarship Data with All Students")
    st.subheader("Upload Files")

    scholarship_file = st.file_uploader("Upload Scholarship_Data.xlsx", type="xlsx")
    original_file = st.file_uploader("Upload Original Excel File", type="xlsx")

    if scholarship_file is not None and original_file is not None:
        scholarship_df = pd.read_excel(scholarship_file)
        original_df = pd.read_excel(original_file)
        combined_df = process_files(scholarship_df, original_df)

        if combined_df is not None:
            st.subheader("Combined Data")
            st.write(combined_df)
            download_combined_df(combined_df)

            summary_df = calculate_summary(combined_df.copy())
            if summary_df is not None:
                st.subheader("Scholarship Summary")
                st.write(summary_df)
                download_summary_df(summary_df)
    else:
        st.info("Please upload both files to proceed.")

def process_files(scholarship_df, original_df):
    st.subheader("Processing Files")

    if 'Neptun kód' not in scholarship_df.columns or 'Neptun kód' not in original_df.columns:
        st.error("Error: 'Neptun kód' column is missing in one of the files.")
        return None

    columns_to_keep = scholarship_df.columns.tolist()

    merged_df = pd.merge(
        original_df,
        scholarship_df[['Neptun kód']],
        on='Neptun kód',
        how='left',
        indicator=True
    )
    new_students_df = merged_df[merged_df['_merge'] == 'left_only'].copy()

    existing_columns = [col for col in columns_to_keep if col in new_students_df.columns]
    new_students_df = new_students_df[existing_columns]

    missing_cols = [col for col in columns_to_keep if col not in new_students_df.columns]

    if 'Kredit szám' in missing_cols:
        if 'ElőzőFélévTeljesítettKredit' in new_students_df.columns:
            new_students_df['Kredit szám'] = new_students_df['ElőzőFélévTeljesítettKredit']
        elif 'ElőzőFélévTeljesítettKredit' in original_df.columns:
            kredits = original_df[['Neptun kód', 'ElőzőFélévTeljesítettKredit']]
            new_students_df = pd.merge(new_students_df, kredits, on='Neptun kód', how='left')
            new_students_df['Kredit szám'] = new_students_df['ElőzőFélévTeljesítettKredit']
        else:
            new_students_df['Kredit szám'] = ''
        missing_cols.remove('Kredit szám')

    for col in missing_cols:
        new_students_df[col] = ''

    new_students_df = new_students_df[columns_to_keep]

    combined_df = pd.concat([scholarship_df, new_students_df], ignore_index=True)

    if 'GroupIndex' not in combined_df.columns:
        st.error("Error: 'GroupIndex' column is missing in the data.")
        return None
    if 'KépzésNév' not in combined_df.columns or \
            'Nyelv ID' not in combined_df.columns or \
            'Képzési szint_x' not in combined_df.columns:
        st.error("Error: Required columns ('KépzésNév', 'Nyelv ID', 'Képzési szint_x') are missing in the data.")
        return None

    group_min_osztondijindex = scholarship_df[['GroupIndex', 'Group Minimum Ösztöndíjindex']].drop_duplicates()
    group_min_osztondijindex = group_min_osztondijindex.dropna(subset=['Group Minimum Ösztöndíjindex'])

    combined_df = pd.merge(
        combined_df.drop(columns=['Group Minimum Ösztöndíjindex']),
        group_min_osztondijindex,
        on='GroupIndex',
        how='left'
    )

    required_columns = ['Ösztöndíj átlag előző félév', 'ElőzőFélévTeljesítettKredit']
    for col in required_columns:
        if col not in combined_df.columns:
            st.error(f"Error: Column '{col}' is missing in the data.")
            return

    combined_df['Ösztöndíj átlag előző félév'] = pd.to_numeric(combined_df['Ösztöndíj átlag előző félév'],
                                                               errors='coerce')
    combined_df['ElőzőFélévTeljesítettKredit'] = pd.to_numeric(combined_df['ElőzőFélévTeljesítettKredit'],
                                                               errors='coerce')

    conditions_met = (
            (combined_df['Ösztöndíj átlag előző félév'] >= 3.8) &
            (combined_df['ElőzőFélévTeljesítettKredit'] >= 23) &
            (combined_df['Exceed Limit'] == False))

    scholarship_amount_idx = combined_df.columns.get_loc('Scholarship Amount')

    combined_df.insert(scholarship_amount_idx, 'Jogosultság döntés',
                       conditions_met.map({True: 'Jogosult', False: 'Nem Jogosult'}))


    def determine_indoklas(row):
        if row['Jogosultság döntés'] == 'Jogosult':
            return 'Ösztöndíjra jogosult'
        else:
            reasons = []
            if row['Exceed Limit'] == True:
                reasons.append('Túllépte a jogosultsági időszakot')
            if row['Ösztöndíj átlag előző félév'] < 3.8:
                reasons.append('Nem érte el a minimum átlagot')
            if row['ElőzőFélévTeljesítettKredit'] < 23:
                reasons.append('Nem érte el a minimum kreditet')
            return ' és '.join(reasons)



    combined_df['Jogosultság indoklás'] = combined_df.apply(determine_indoklas, axis=1)

    jogosultsag_dontes_idx = combined_df.columns.get_loc('Jogosultság döntés')
    combined_df.insert(jogosultsag_dontes_idx + 1, 'Jogosultság indoklás', combined_df.pop('Jogosultság indoklás'))
    jogosultsag_indoklas_idx = combined_df.columns.get_loc('Jogosultság indoklás')
    combined_df['Scholarship Amount'] = pd.to_numeric(combined_df['Scholarship Amount'],
                                                               errors='coerce')
    combined_df['Scholarship Amount'] = pd.to_numeric(combined_df['Scholarship Amount'], errors='coerce')
    combined_df['Scholarship Amount'] = (combined_df['Scholarship Amount'] / 100).round() * 100

    conditions_met2 = (combined_df['Scholarship Amount'] > 1)
    combined_df.insert(jogosultsag_indoklas_idx + 1, 'Ösztöndíj döntés',
                       conditions_met2.map({True: 'Jogosult', False: 'Nem Jogosult'}))

    combined_df['Group Minimum Ösztöndíjindex'] = pd.to_numeric(combined_df['Group Minimum Ösztöndíjindex'],
                                                               errors='coerce')
    combined_df['Ösztöndíjindex'] = pd.to_numeric(combined_df['Ösztöndíjindex'], errors='coerce')

    def determine_osztondij_indoklas(row):
        if pd.isna(row['Hallgató kérvény azonosító']):
            return ' '
        elif row['Ösztöndíj döntés'] == 'Jogosult':
            return ('Ösztöndíjt kap')
        elif row['Exceed Limit']:
            return 'Túllépte a jogosultsági időszakot'
        elif row['Ösztöndíj átlag előző félév'] < 3.8 and row['ElőzőFélévTeljesítettKredit'] < 23:
            return("Nem érte el a szükséges átlagot és kreditet")
        elif row['Ösztöndíj átlag előző félév'] < 3.8:
            return ('Nem érte el a szükséges átlagot')
        elif row['ElőzőFélévTeljesítettKredit'] < 23:
            return ('Nem érte el a szükséges kreditet')
        elif row['Group Minimum Ösztöndíjindex'] > row['Ösztöndíjindex']:
            return ('Nem érte el a csoportja minimum ösztöndíjindexét')
        else:
            return 'Egyéb ok'

    combined_df['Ösztöndíj indoklás'] = combined_df.apply(determine_osztondij_indoklas, axis=1)
    osztondij_dontes_idx = combined_df.columns.get_loc('Ösztöndíj döntés')
    combined_df.insert(osztondij_dontes_idx + 1, 'Ösztöndíj indoklás', combined_df.pop('Ösztöndíj indoklás'))

    negy_havi_osztondij_idx = combined_df.columns.get_loc('Scholarship Amount') + 1
    combined_df.insert(negy_havi_osztondij_idx, '5 havi Ösztöndíj', combined_df['Scholarship Amount'] * 5)

    combined_df = combined_df.rename(columns={'Scholarship Amount' : '1 havi Ösztöndíj'})
    combined_df.insert(0, 'ID', range(1, len(combined_df) + 1))
    return combined_df

def download_combined_df(combined_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        combined_df.to_excel(writer, index=False, sheet_name='Combined_Data')
        workbook = writer.book
        worksheet = writer.sheets['Combined_Data']

        last_row = combined_df.shape[0]
        last_col = combined_df.shape[1]

        yellow_fill = workbook.add_format({'bg_color': '#FFFF00'})
        group_fill_1 = workbook.add_format({'bg_color': '#D3D3D3'})
        group_fill_2 = workbook.add_format({'bg_color': '#FFFFFF'})

        col_idx = combined_df.columns.get_loc('1 havi Ösztöndíj')
        worksheet.conditional_format(1, col_idx, last_row, col_idx, {'type': 'no_blanks', 'format': yellow_fill})

        combined_df['TempGroupID'] = combined_df[['GroupIndex']].apply(lambda x: ' | '.join(x.astype(str)), axis=1)
        previous_group = None
        current_fill = group_fill_1

        for row in range(1, last_row + 1):
            group_id = combined_df.iloc[row - 1]['TempGroupID']

            if group_id != previous_group:
                current_fill = group_fill_1 if current_fill == group_fill_2 else group_fill_2
                previous_group = group_id

            worksheet.set_row(row, None, current_fill)

        combined_df.drop(columns=['TempGroupID'], inplace=True)
    output.seek(0)

    st.download_button(
        label='Download Combined Excel File',
        data=output.getvalue(),
        file_name='Combined_Scholarship_Data.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

def download_summary_df(summary_df):
    """Downloads the summary DataFrame as an Excel file."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        summary_df.to_excel(writer, sheet_name='Summary_Data', index=True)
    output.seek(0)

    st.download_button(
        label='Download Summary Excel File',
        data=output.getvalue(),
        file_name='Scholarship_Summary.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

if __name__ == "__main__":
    main()

