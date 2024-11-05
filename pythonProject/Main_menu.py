import streamlit as st

def main():

    st.title("Scholarship Calculation Main Menu")
    st.subheader("This is the main page of the application.", divider=True)
    st.markdown("### Navigate to Other Pages")

    with st.container(border=True):
        st.subheader("Step 1 - _Grouping of students_")
        st.markdown("[Go to page](https://studentgrouping.streamlit.app)")

    with st.container(border=True):
        st.subheader("Step 2 - _Calculation of scholarship_")
        st.markdown("[Go to page](https://scholarshipcalc.streamlit.app)")

    with st.container(border=True):
        st.subheader("Step 3 - _Final merge of excels_")
        st.markdown("[Go to page](-)")

if __name__ == "__main__":
    main()