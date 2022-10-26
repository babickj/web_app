
# Import required libraries
# from tkinter import Toplevel
# import PyPDF2
# import textract
import re
import string
import pandas as pd
# import matplotlib.pyplot as plt
from os import path

from datetime import datetime
from datetime import date
# import tkinter as tk
# from tkcalendar import Calendar

# import seaborn as sns
# import matplotlib.pyplot as plt
import numpy as np

from glob import glob

# import job_dict
import sasm_globals

SUSP_ABS = sasm_globals.SUSP_ABS
FT2_GOAL = sasm_globals.FT2_GOAL
DL_ABS_VAR_ITM_target = sasm_globals.DL_ABS_VAR_ITM_target
Billability_Target = sasm_globals.Billability_Target
ALLOWED_VAR = sasm_globals.ALLOWED_VAR

def pre_process(mbt):

     # Split out the employee ID from the employee name, this allows us to merge the labor and the bench data
    mbt.df_bench[['Full Name', 'Emplid']] = mbt.df_bench['Empl Name & ID'].str.split("(", expand=True)
    mbt.df_bench['Emplid'] = mbt.df_bench['Emplid'].str.replace(r'\D', '', regex=True)

    # Make a column to show on original bench report, this will help in later data Analysis
    mbt.df_bench['on_bench'] = True
    mbt.df_bench['Notes'] = "On official bench report"

    # Merge the bench and labor data for easier analysis
    # Set index for both dataframes
    mbt.df_consolidated = mbt.df_labor.set_index('Emplid').join(mbt.df_bench.set_index('Emplid'), rsuffix='_bench', how='outer')

    # Drop the 'Grand Total' row
    mbt.df_consolidated = mbt.df_consolidated.drop('Grand Total')

    # Updated index to int
    mbt.df_consolidated.index = mbt.df_consolidated.index.astype(int)

    # If employee is not on labor report, transfer his name for later updating
    mbt.df_consolidated["Empl Full Name "].fillna(mbt.df_consolidated["Full Name"], inplace = True)
    mbt.df_consolidated["Empl Name & ID"].fillna(mbt.df_consolidated["Empl Full Name "], inplace = True)

    # Rename Emp ID column on last SASM report to facilitate joinging
    mbt.df_last_sasm = mbt.df_last_sasm.rename(columns={'Emp ID':'Emplid'})
    
    # Join the last sasm to consolidated
    mbt.df_last_sasm =  mbt.df_last_sasm.set_index('Emplid')
 
    # Create new empty columns to transfer data from the last sasm report
    mbt.df_consolidated["SASM Notes"] = ""
    mbt.df_consolidated["SASM RM Status"] = ""
    mbt.df_consolidated["Anticipated Start Date"] = ""
    mbt.df_consolidated["Activity Comments"] = ""
    
    # mbt.df_consolidated = mbt.df_consolidated.drop_duplicates(keep='first')
    mbt.df_consolidated  = mbt.df_consolidated [~mbt.df_consolidated.index.duplicated(keep='first')]
    # This loop moves specific cells from one .xlxs file to the dataframe.
    for idx in mbt.df_consolidated.index:
        
        for idx_2 in mbt.df_last_sasm.index:

            # If employee had notes made on him/her last week, copy
            # that data over
            if idx == idx_2:
                
                tmp_notes = mbt.df_last_sasm.loc[mbt.df_last_sasm.index == idx_2 ,'SASM Notes']
                tmp_status = mbt.df_last_sasm.loc[mbt.df_last_sasm.index == idx_2 ,'SASM RM Status']
                tmp_start = mbt.df_last_sasm.loc[mbt.df_last_sasm.index == idx_2 ,'Anticipated Start Date']
                tmp_comment = mbt.df_last_sasm.loc[mbt.df_last_sasm.index == idx_2 ,'Activity Comments']
                mbt.df_consolidated.loc[mbt.df_consolidated.index == idx, "SASM Notes"] = tmp_notes
                mbt.df_consolidated.loc[mbt.df_consolidated.index == idx, "SASM RM Status"] = tmp_status
                mbt.df_consolidated.loc[mbt.df_consolidated.index == idx, "Anticipated Start Date"] = tmp_start
                mbt.df_consolidated.loc[mbt.df_consolidated.index == idx, "Activity Comments"] = tmp_comment
    
    
    
    # Set all labor report additions to False
    mbt.df_consolidated.loc[mbt.df_consolidated['on_bench'] != True, "on_bench"] = False
    mbt.df_consolidated['on_bench'] = mbt.df_consolidated['on_bench'].fillna(False)

    # Add data from the previous bench report
    mbt.df_consolidated['prev_bench'] = False
    
    # Process previous bench data:
    # Split out the employee ID from the employee name, this allows us to merge the labor and the bench data
    mbt.df_last_bench[['Full Name', 'Emplid']] = mbt.df_last_bench['Empl Name & ID'].str.split("(", expand=True)
    mbt.df_last_bench['Emplid'] = mbt.df_bench['Emplid'].str.replace(r'\D', '', regex=True)
    mbt.df_last_bench = mbt.df_last_bench.set_index('Emplid')
    mbt.df_last_bench.index = mbt.df_last_bench.index.fillna(0)
    mbt.df_last_bench.index = mbt.df_last_bench.index.map(int)
    
    # Annotate 2 reporting cycles on main dataframe
    mbt.df_consolidated.loc[mbt.df_consolidated.index.isin(mbt.df_last_bench.index),'prev_bench'] = True

    mbt.df_consolidated['Billability Variance to Target'] = mbt.df_consolidated['Billability Variance to Target'].str.replace('%','',regex=True)
    mbt.df_consolidated['Billability Variance to Target'] = mbt.df_consolidated['Billability Variance to Target'].fillna(0)
    mbt.df_consolidated['Billability Variance to Target'] = mbt.df_consolidated['Billability Variance to Target'].map(float)

    mbt.df_consolidated['Billability'] = mbt.df_consolidated['Billability'].str.replace('%','',regex=True)
    mbt.df_consolidated['Billability'] = mbt.df_consolidated['Billability'].map(float)

    mbt.df_consolidated['Billability Target'] = mbt.df_consolidated['Billability Target'].fillna(0)
    mbt.df_consolidated['Billability Target'] = mbt.df_consolidated['Billability Target'].str.replace('%','',regex=True)
    mbt.df_consolidated['Billability Target'] = mbt.df_consolidated['Billability Target'].map(float)

    ### PROCESS LAST WATCH LIST DATA #########
    if mbt.df_last_watch != None:
        mbt.df_last_watch.set_index('Empl ID')
        mbt.df_last_watch.index = mbt.df_last_watch.index.map(int)
        
        
def dt_less_4_bd(mbt, user):

    # Get date/time of file
    try:
      
        f_date = user.bench_date
        if f_date == "No File":
            f_date = datetime.fromtimestamp(path.getmtime(user.bench_path))


    except Exception as e:
        # print(f"{e}: INVALID BENCH FILE")
        f_date = datetime.today()


    # Set dates for checking
    today = date.today()
    f_date = datetime.date(f_date)
    start_one = date(today.year,today.month,sasm_globals.periods['1st'][0])
    start_two = date(today.year,today.month,sasm_globals.periods['2nd'][0])
    p1_lenth = sasm_globals.periods['1st'][1] - sasm_globals.periods['1st'][0]
    p2_lenth = sasm_globals.periods['2nd'][1] - sasm_globals.periods['2nd'][0]

    # Check if with periods
    if np.busday_count(start_one, f_date) <= p1_lenth:
        return True
    elif (f_date >= start_two) and (np.busday_count(start_two, f_date) <= p2_lenth):
        return True

    return False

def dt_isBillable(mbt):
    
    # print(mbt.df_consolidated.loc[mbt.df_consolidated.index==577354]['on_bench'])
    # print(mbt.df_consolidated.loc[mbt.df_consolidated.index==577354]['Billability Variance to Target'])
    # print(mbt.df_consolidated.loc[mbt.df_consolidated.index==577354]['Billability Target'] * ALLOWED_VAR)

    # Check to see if billable
    mbt.df_consolidated.loc[(mbt.df_consolidated['Billability Variance to Target'] > (mbt.df_consolidated['Billability Target'] * ALLOWED_VAR))
        & (mbt.df_consolidated['on_bench'] == True),
        "Notes"] = "Billable - Removed from Bench"

    mbt.df_consolidated.loc[(mbt.df_consolidated['Billability Variance to Target'] > (mbt.df_consolidated['Billability Target'] * ALLOWED_VAR))
        & (mbt.df_consolidated['on_bench'] == True),
        "on_bench"] = False

# (SUSP + ABS > 0.75) and (NOT on prior period bench)
def dt_pt_2(mbt):

        # Check to see if suspended AND ABS > 0.75.  We will remove people from bench if this is true
        # Convert strings to int for data manipulation
        mbt.df_consolidated['Suspense Amount'] = mbt.df_consolidated['Suspense Amount'].fillna(0)
        mbt.df_consolidated['Suspense Amount'] = mbt.df_consolidated['Suspense Amount'].replace({'\$':'',',':'','\(':'-','\)':''},regex=True).astype(float)

        mbt.df_consolidated['DL $ Target '] = mbt.df_consolidated['DL $ Target '].replace({'\(':'-', '\)':'',',':''}, regex=True)
        mbt.df_consolidated['DL $ Target '] = mbt.df_consolidated['DL $ Target '].map(float)

        mbt.df_consolidated.loc[(mbt.df_consolidated['on_bench'] == True) 
            & (mbt.df_consolidated['Suspense Amount'] >= (mbt.df_consolidated['DL $ Target ']*SUSP_ABS))
            & (mbt.df_consolidated['prev_bench'] != True),  
            'Notes'] = f"Removed - SUSP > {SUSP_ABS} and NOT on previous bench."
        
        mbt.df_consolidated.loc[(mbt.df_consolidated['on_bench'] == True) 
            & (mbt.df_consolidated['Suspense Amount'] >= (mbt.df_consolidated['DL $ Target ']*SUSP_ABS))
            & (mbt.df_consolidated['prev_bench'] != True),
             'on_bench'] = False   

        ###
        mbt.df_consolidated.loc[(mbt.df_consolidated['on_bench'] == True) 
            & (mbt.df_consolidated['Suspense Amount'] >= (mbt.df_consolidated['DL $ Target ']*SUSP_ABS))
            & (mbt.df_consolidated['prev_bench'] == True),  
            'Notes'] = f"Watch List - SUSP > {SUSP_ABS} and on previous bench."
        
        mbt.df_consolidated.loc[(mbt.df_consolidated['on_bench'] == True) 
            & (mbt.df_consolidated['Suspense Amount'] >= (mbt.df_consolidated['DL $ Target ']*SUSP_ABS))
            & (mbt.df_consolidated['prev_bench'] == True),
             'on_bench'] = 'watch'   

        # Check to see is SUSP + ABS make up for the difference
        mbt.df_consolidated['Total Absence Amount'] = mbt.df_consolidated['Total Absence Amount'].replace({'\$':'','\(':'-','\)':'',',':''}, regex=True)
        mbt.df_consolidated['Total Absence Amount'] = mbt.df_consolidated['Total Absence Amount'].map(float)
        mbt.df_consolidated.loc[
            # check suspense amount
            (((mbt.df_consolidated['Suspense Amount'] / mbt.df_consolidated['DL $ Target ']) + 
            (mbt.df_consolidated['Total Absence Amount'] / mbt.df_consolidated['DL $ Target '])) > SUSP_ABS)
            & (mbt.df_consolidated['on_bench'] == True)
            & (mbt.df_consolidated['prev_bench'] != True),
            'Notes'] = f"Removed - [SUSP + ABS] > {SUSP_ABS}, and NOT on previous bench."

        mbt.df_consolidated.loc[
            # check suspense amount
            (((mbt.df_consolidated['Suspense Amount'] / mbt.df_consolidated['DL $ Target ']) + 
            (mbt.df_consolidated['Total Absence Amount'] / mbt.df_consolidated['DL $ Target '])) > SUSP_ABS)
            & (mbt.df_consolidated['on_bench'] == True)
            & (mbt.df_consolidated['prev_bench'] != True),
            'on_bench'] = False

        # Check to see if FT2 is high
        # Process FT2 data
        mbt.df_consolidated['+2 Month FT2 DL Hrs'] = mbt.df_consolidated['+2 Month FT2 DL Hrs'].fillna(0)
        mbt.df_consolidated['+2 Month FT2 DL Hrs'] = mbt.df_consolidated['+2 Month FT2 DL Hrs'].map(int) 

        mbt.df_consolidated.loc[
            (mbt.df_consolidated['on_bench'] == True) 
            & (((mbt.df_consolidated['Suspense Amount'] / mbt.df_consolidated['DL $ Target ']) + 
            (mbt.df_consolidated['Total Absence Amount'] / mbt.df_consolidated['DL $ Target '])) > SUSP_ABS)
            & (mbt.df_consolidated['+2 Month FT2 DL Hrs'] >= FT2_GOAL),
            'Notes'] = f"Removed - FT2 > {FT2_GOAL} with high suspense."

        mbt.df_consolidated.loc[
            (mbt.df_consolidated['on_bench'] == True) 
            & (((mbt.df_consolidated['Suspense Amount'] / mbt.df_consolidated['DL $ Target ']) + 
            (mbt.df_consolidated['Total Absence Amount'] / mbt.df_consolidated['DL $ Target '])) > SUSP_ABS)
            & (mbt.df_consolidated['Suspense Amount'] > SUSP_ABS)
            & (mbt.df_consolidated['+2 Month FT2 DL Hrs'] >= FT2_GOAL),
            'on_bench'] = False

        # If high-suspense, and on previous bench, place on watchlist
        mbt.df_consolidated.loc[
            # check suspense amount
            (((mbt.df_consolidated['Suspense Amount'] / mbt.df_consolidated['DL $ Target ']) + 
            (mbt.df_consolidated['Total Absence Amount'] / mbt.df_consolidated['DL $ Target '])) > SUSP_ABS)
            & (mbt.df_consolidated['on_bench'] == True)
            & (mbt.df_consolidated['prev_bench'] != True),
            'Notes'] = f"Watch List - [SUSP + ABS] > {SUSP_ABS}."

        mbt.df_consolidated.loc[
            # check suspense amount
            (((mbt.df_consolidated['Suspense Amount'] / mbt.df_consolidated['DL $ Target ']) + 
            (mbt.df_consolidated['Total Absence Amount'] / mbt.df_consolidated['DL $ Target '])) > SUSP_ABS)
            & (mbt.df_consolidated['on_bench'] == True)
            & (mbt.df_consolidated['prev_bench'] != True),
            'on_bench'] = "watch"

        
def dt_pt_3(mbt):

    # Pre-process data:
    mbt.df_consolidated['DL+Absc Hrs Variance to ITM Target %'] = mbt.df_consolidated['DL+Absc Hrs Variance to ITM Target %'].fillna(0)
    mbt.df_consolidated['DL+Absc Hrs Variance to ITM Target %'] = mbt.df_consolidated['DL+Absc Hrs Variance to ITM Target %'].str.rstrip("%")
    mbt.df_consolidated['DL+Absc Hrs Variance to ITM Target %'] = mbt.df_consolidated['DL+Absc Hrs Variance to ITM Target %'].map(float)
    
    # (DL + ABS HRS VAR to ITM target % < -0.50) && (Billability Target > 0.50)?
    mbt.df_consolidated.loc[
        (mbt.df_consolidated['DL+Absc Hrs Variance to ITM Target %'] < DL_ABS_VAR_ITM_target)
        & (mbt.df_consolidated['Billability Target'] > Billability_Target)
        & (mbt.df_consolidated['on_bench'] != True),
        # & (mbt.df_consolidated['prev_bench'] != True),
        "Notes"] = f"Watch list - DL+Absc Hrs Variance to ITM Target % < {DL_ABS_VAR_ITM_target}"

    mbt.df_consolidated.loc[
        (mbt.df_consolidated['DL+Absc Hrs Variance to ITM Target %'] < DL_ABS_VAR_ITM_target)
        & (mbt.df_consolidated['Billability Target'] > Billability_Target)
        & (mbt.df_consolidated['on_bench'] != True),
        # & (mbt.df_consolidated['prev_bench'] != True),
        "on_bench"] = "watch"

    # If on previous watchlist, and still meet criteria, keep on watchlist
    # mbt.df_consolidated['prev_watch'] == False
    # mbt.df_consolidated.loc[(mbt.df_consolidated.index.isin(mbt.df_last_watch.index))
    #                         & (mbt.df_consolidated['DL+Absc Hrs Variance to ITM Target %'] > DL_ABS_VAR_ITM_target)
    #                         , 'on_bench'] = "watch"

    
def dt_pt_4(mbt):
    
    # If employee was on last official SASM, don't move to watchlist
    mbt.df_consolidated.loc[(mbt.df_consolidated.index.isin(mbt.df_last_sasm.index))
                            & (mbt.df_consolidated['on_bench'] == 'watch'),
                            'Notes'] = "Kept on SASM - Was on last official SASM and NOT moved to watchlist"


    # If employee was on last official SASM, don't move to watchlist
    mbt.df_consolidated.loc[(mbt.df_consolidated.index.isin(mbt.df_last_sasm.index))
                            & (mbt.df_consolidated['on_bench'] == 'watch'),
                            'on_bench'] = True

    
def select_date(window,user,date_label):
    
    new_window = Toplevel(window)
    new_window.title("Calendar")
    new_window.geometry("200x200")
    cal = Calendar(new_window, selectmode='day',
                    year = user.bench_date.year,
                    month = user.bench_date.month,
                    day = user.bench_date.day)
    # cal.frame(new_window)
    cal.grid(row=0, column=0)

    def new_date():
        new_window.destroy()
        # print(f"cal: {cal.get_date()}")
        user.bench_date = datetime.strptime(cal.get_date(), '%m/%d/%y')
        tmp_date = f"{user.bench_date.month}/{user.bench_date.day}/{user.bench_date.year}"
        date_label['text'] = tmp_date
        # user.bench_date = tmp_date
        # window.
        # date_label[]

        
    btn_select = tk.Button(new_window, text="Update Date", command=new_date)
    btn_select.grid(row=1, column=0)


def get_files_from_dir(my_path, ext, ig_case):

    if ig_case:
        ext =  "".join(["[{}]".format(ch + ch.swapcase()) for ch in ext])
    return glob(path.join(my_path, "*." + ext))


def deconstruct_file(file_path):
    # Open pdf file
    pdfFileObj = open(file_path,'rb')

    # Read file
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj, strict=False)

    # Get total number of pages
    num_pages = pdfReader.numPages

    # Initialize a count for the number of pages
    count = 0

    # Initialize a text empty etring variable
    text = ""

    # Extract text from every page on the file
    while count < num_pages:
        pageObj = pdfReader.getPage(count)
        count +=1
        text += pageObj.extractText()

    return text


def clean_text(text):
    # Convert all strings to lowercase
    text = text.lower()

    # Remove numbers
    text = re.sub(r'\d+','',text)

    # Remove punctuation
    text = text.translate(str.maketrans('','',string.punctuation))

    return text


def calculate_scores(text):
    # Initializie score counters for each area
    quality = 0
    operations = 0
    supplychain = 0
    project = 0
    data = 0
    healthcare = 0

    # Create an empty list where the scores will be stored
    scores = []

    terms = job_dict.terms

    # Obtain the scores for each area
    for area in terms.keys():
            
        if area == 'Quality/Six Sigma':
            for word in terms[area]:
                if word in text:
                    quality +=1
            scores.append(quality)
            
        elif area == 'Operations management':
            for word in terms[area]:
                if word in text:
                    operations +=1
            scores.append(operations)
            
        elif area == 'Supply chain':
            for word in terms[area]:
                if word in text:
                    supplychain +=1
            scores.append(supplychain)
            
        elif area == 'Project management':
            for word in terms[area]:
                if word in text:
                    project +=1
            scores.append(project)
            
        elif area == 'Data analytics':
            for word in terms[area]:
                if word in text:
                    data +=1
            scores.append(data)
            
        else:
            for word in terms[area]:
                if word in text:
                    healthcare +=1
            scores.append(healthcare)

        # Create a data frame with the scores summary
        
        # print(terms.keys())
    # summary = pd.DataFrame(scores,index=terms.keys(),columns=['score']).sort_values(by='score',ascending=False)
    summary = pd.DataFrame(data=scores, columns=['score'], index=terms.keys()).sort_values(by='score', ascending=False)
   
    return summary
    
def get_data_from_dir(dir_path, ext, ig_case):
    # For all files, construct a score and add a column
    # to the employee df
    df_emp = pd.DataFrame()
    applicant_list = {}
    app_num = 0

    files = get_files_from_dir(dir_path, ext, ig_case)

    # Loops through all the files in the given directory
    # and puts the employee scores into a dataframe
    for file in files:

        text = deconstruct_file(file)
        text = clean_text(text)

        if df_emp.empty:
            df_emp = calculate_scores(text)
        else:
            df_emp = pd.concat([df_emp, calculate_scores(text)], axis=1)
        
        df_emp.rename(columns={'score':f"_{app_num}"}, inplace=True)
        # Add applicant to dic
        applicant_list[file] = app_num

        # Increment number of applicants
        app_num += 1

   

    return df_emp, applicant_list

def match_employees(df_emp, df_jobs):

    # Get all the skills catagories
    categories = df_emp.index

    # List of employees by number
    empl = df_emp.columns

    # List of jobs by number
    jobs = df_jobs.columns
    
    # Create an empty dataframe to import the skill match to
    df_match=pd.DataFrame(index=df_jobs.columns, columns=df_emp.columns)
    
    # Create a master df for easier analysis
    df_tmp = pd.concat([df_emp, df_jobs], axis = 1)

    # Loop through all employees and all jobs
    # Calculate fit score for each pairing
    for emp in empl:
        tmp = 0
        for job in jobs:
            tmp = sum(df_tmp[emp]- df_tmp[job])
            df_match.loc[job][emp] = tmp
    
    # return the df of fit
    return df_match
    
def display_job_recommendations(df_match):

    # Hold list of employees
    empl = df_match.columns
    jobs = df_match.index

    match = {}

    # Print recommendations for each employee
    for emp in empl:

        match[emp] = []
        for job in jobs:

            if(df_match.loc[job, emp] > 0):
                match[emp].append(job)
    
    # print(match)

    fig, ax = plt.subplots()
    sns.heatmap(df_match,ax=ax, cmap="YlGnBu")
    plt.show()

    return match
