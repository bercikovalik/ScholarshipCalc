import streamlit as st
import pandas as pd
from io import BytesIO

def main():
    st.title("Final Step: Merge Scholarship Data with All Students")
    st.subheader("Upload Files")

    scholarship_file = st.file_uploader("Upload Scholarship_Data.xlsx", type="xlsx")
    original_file = st.file_uploader("Upload Original Excel File", type="xlsx")

    if scholarship_file is not None and original_file is not None:
        scholarship_df = pd.read_excel(scholarship_file)

        original_df = pd.read_excel(original_file)

        process_files(scholarship_df, original_df)
    else:
        st.info("Please upload both files to proceed.")

def process_files(scholarship_df, original_df):
    st.subheader("Processing Files")

    if 'Neptun kód' not in scholarship_df.columns or 'Neptun kód' not in original_df.columns:
        st.error("Error: 'Neptun kód' column is missing in one of the files.")
        return

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
        return

    group_min_kodi = scholarship_df[['GroupIndex', 'Group Minimum KÖDI']].drop_duplicates()
    group_min_kodi = group_min_kodi.dropna(subset=['Group Minimum KÖDI'])

    combined_df = pd.merge(
        combined_df.drop(columns=['Group Minimum KÖDI']),
        group_min_kodi,
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

    conditions_met = (combined_df['Ösztöndíj átlag előző félév'] >= 3.8) & (
                combined_df['ElőzőFélévTeljesítettKredit'] >= 23)

    scholarship_amount_idx = combined_df.columns.get_loc('Scholarship Amount')

    combined_df.insert(scholarship_amount_idx, 'Jogosultság döntés',
                       conditions_met.map({True: 'Jogosult', False: 'Nem Jogosult'}))

    def determine_indoklas(row):
        if row['Jogosultság döntés'] == 'Jogosult':
            return 'Ösztöndíjra jogosult'
        else:
            reasons = []
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
    conditions_met2 = (combined_df['Scholarship Amount'] > 1)
    combined_df.insert(jogosultsag_indoklas_idx + 1, 'Ösztöndíj döntés',
                       conditions_met2.map({True: 'Jogosult', False: 'Nem Jogosult'}))

    combined_df['Group Minimum KÖDI'] = pd.to_numeric(combined_df['Group Minimum KÖDI'],
                                                               errors='coerce')
    combined_df['KÖDI'] = pd.to_numeric(combined_df['KÖDI'], errors='coerce')

    def determine_osztondij_indoklas(row):
        if pd.isna(row['Hallgató kérvény azonosító']):
            return 'Nem pályázott'
        elif row['Ösztöndíj döntés'] == 'Jogosult':
            return 'Jogosult'
        elif row['Ösztöndíj átlag előző félév'] < 3.8:
            return 'Nem érte el a minimum átlagot'
        elif row['ElőzőFélévTeljesítettKredit'] < 23:
            return 'Nem érte el a minimum kreditet'
        elif row['Group Minimum KÖDI'] > row['KÖDI']:
            return 'Nem érte el a csoport minimum átlagát'
        else:
            return 'Egyéb ok'

    combined_df['Ösztöndíj indoklás'] = combined_df.apply(determine_osztondij_indoklas, axis=1)
    osztondij_dontes_idx = combined_df.columns.get_loc('Ösztöndíj döntés')
    combined_df.insert(osztondij_dontes_idx + 1, 'Ösztöndíj indoklás', combined_df.pop('Ösztöndíj indoklás'))

    st.subheader("Combined Data")
    st.write(combined_df)

    download_combined_df(combined_df)

def download_combined_df(combined_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        combined_df.to_excel(writer, index=False, sheet_name='Combined_Data')

    output.seek(0)

    st.download_button(
        label='Download Combined Excel File',
        data=output.getvalue(),
        file_name='Combined_Scholarship_Data.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == "__main__":
    main()
