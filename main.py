import streamlit as st
import os
import SASM_v4 as sa
import pandas as pd
# from io import BytesIO
from io import StringIO
from file_process import import_bench
from file_process import import_last_bench
from file_process import import_labor_data
from file_process import side_bar
from file_process import import_last_sasm

def main(user, mbt):
    """
    Purpose: Sets the layout of the web app
    Params:  None
    Return:  None
    """

    
    # Display the header / welcome
    st.set_page_config(page_title='Bench Report Generator', layout="wide")
    st.title("_SIG Analytics Staff Management_")
    
    ##############################################
    #
    # CREATE CENTER OF PAGE
    #
    ##############################################
    with st.container():
        h1, h2 = st.columns(2, gap="small")
        with h1:
            st.header("**Status Tracker**")
        with h2:
            st.header("**Import Data**")
            
    # Asks user to select file
    with st.container():
    
        # Format page to three colums
        col0, col1, col2, col3 = st.columns(4,gap="small")
        
        # Format col1
        with col1:
            import_bench(user, mbt)
            import_last_sasm(user, mbt)

        # Format col2
        with col2:
            import_last_bench(user,mbt)
    
        with col3:
            import_labor_data(user,mbt)

    ##############################################
    #
    # CREATE SIDE COLUMN
    #
    ##############################################
        with col0:
            side_bar(user, mbt)
            
    ##############################################
    #
    # CREATE DISPLAY TABS
    #
    ##############################################
    # Create tabs to display data to user
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Current Bench", "Last Bench",
                                "Labor Data", "Last SASM Report",
                                "Generated Bench Report"])
    with tab1:
        st.header("Current Bench Data")
        
        if user.bench_path == "No File Selected":
            st.write(user.bench_path)
        else:
            st.write(mbt.df_bench)
                

    with tab2:
        st.header("Last Bench Data")
        
        if user.last_bench == "No File Selected":
            st.write(user.last_bench)
        else:
            st.write(mbt.df_last_bench)

    with tab3:
        st.header("Labor Data")
        if user.labor_path == "No File Selected":
            st.write(user.labor_path)
        else:
            st.write(mbt.df_labor)
            
    with tab4:
        st.header("Last SASM Report")
        if user.last_sasm == "No File Selected":
            st.write(user.last_sasm)
        else:
            st.write(mbt.df_last_sasm)
            
    with tab5:
        st.header("Generated Bench Report")
        if mbt.df_final_report.empty:
            st.write("Upload Date and Generate Report")
        else:
            st.write(mbt.df_final_report)

  

def file_selector(folder_path='.'):
    filenames = os.listdir(folder_path)
    selected_filename = st.selectbox('Most recent bench.csv file', filenames)
    
    file = st.file_uploader("Please choose a file")
    
    return os.path.join(selected_filename)

if __name__ == "__main__":
    # Initialize objects
    
    if 'user' not in st.session_state:
        st.session_state.user = sa.User_Data()
        user = st.session_state.user
    else:
        user = st.session_state.user
    
    if 'mbt' not in st.session_state:
        st.session_state.mbt = sa.Billable_Hour_Tracker()
        mbt = st.session_state.mbt
    else:
        mbt = st.session_state.mbt
        
    
    main(user, mbt)
