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

    # Ensure 'Neptun kód' is present in both DataFrames
    if 'Neptun kód' not in scholarship_df.columns or 'Neptun kód' not in original_df.columns:
        st.error("Error: 'Neptun kód' column is missing in one of the files.")
        return

    # Get the list of columns from Scholarship_Data.xlsx
    columns_to_keep = scholarship_df.columns.tolist()

    # Find students in original_df not present in scholarship_df based on 'Neptun kód'
    merged_df = pd.merge(
        original_df,
        scholarship_df[['Neptun kód']],
        on='Neptun kód',
        how='left',
        indicator=True
    )
    new_students_df = merged_df[merged_df['_merge'] == 'left_only'].copy()

    # Keep only the columns present in Scholarship_Data.xlsx and original_df
    existing_columns = [col for col in columns_to_keep if col in new_students_df.columns]
    new_students_df = new_students_df[existing_columns]

    # Identify missing columns
    missing_cols = [col for col in columns_to_keep if col not in new_students_df.columns]

    # For 'Kredit szám', set it equal to 'ElőzőFélévTeljesítettKredit' if available
    if 'Kredit szám' in missing_cols:
        if 'ElőzőFélévTeljesítettKredit' in new_students_df.columns:
            new_students_df['Kredit szám'] = new_students_df['ElőzőFélévTeljesítettKredit']
        elif 'ElőzőFélévTeljesítettKredit' in original_df.columns:
            # Map 'ElőzőFélévTeljesítettKredit' from original_df
            kredits = original_df[['Neptun kód', 'ElőzőFélévTeljesítettKredit']]
            new_students_df = pd.merge(new_students_df, kredits, on='Neptun kód', how='left')
            new_students_df['Kredit szám'] = new_students_df['ElőzőFélévTeljesítettKredit']
        else:
            new_students_df['Kredit szám'] = ''
        missing_cols.remove('Kredit szám')

    # Add other missing columns with empty strings
    for col in missing_cols:
        new_students_df[col] = ''

    # Reorder columns to match columns_to_keep
    new_students_df = new_students_df[columns_to_keep]

    # Combine scholarship_df with new_students_df
    combined_df = pd.concat([scholarship_df, new_students_df], ignore_index=True)

    # Display the combined DataFrame
    st.subheader("Combined Data")
    st.write(combined_df)

    # Provide an option to download the combined DataFrame
    download_combined_df(combined_df)

def download_combined_df(combined_df):
    # Convert DataFrame to Excel in memory
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

