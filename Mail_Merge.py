""" This script works as a macro to organize the Course Report for start of term. 
    It gets data from four files containing information about courses, program blocks, program contacts,
    and program coordinators, and combines the relevant information to produce a csv intended for 
    Math Consultants to use to contact course personnel and schedule class visits"""

## Imports:
import pandas as pd
import numpy as np
import xlrd
import os
import openpyxl

import tkinter as Tk
from tkinter import filedialog

## ----------------- Functions --------------------------------------

#Given the name of a coordinator, find their email address from df_pgrm_coord
def find_email(name):
    """
        Given the name of a coordinator, find their email address from df_pgrm_coord
    """
    try:
        return df_pgrm_coord[df_pgrm_coord["Textbox28"] == name]["Textbox32"].iloc[0]
    except:
        return "No Match"

#Given a list of names, find the email addresses for the coordinators in a dictionary
def get_emails(names):
    """ 
        Given a list of names, find the email addresses for the coordinators in a dictionary
    """
    if names == ['No Match'] or not names:
        return {'No Match': 'No Match'}
    return {name: find_email(name) for name in names}

#Given a dictionary of {codes:names}, get all coordinator emails associated with that code
def get_all_emails(dnames):
    """ 
        Given a dictionary of {codes:names}, get all coordinator emails associated with that code
    """
    return {code: [get_prg_name(code), get_emails(dnames[code])] for code in dnames.keys()}

#Given a program code, find the list of names for coordinators associated with that code from df_cntct_lst
def get_names(code):
    """
        Given a program code, find the list of names for coordinators associated with that code from df_cntct_lst
    """
    if code == 'No Match':
        return ['No Match']
    return list(filter(None, list(df_cntct_lst[df_cntct_lst['Program Code'] == code]['Coordinator'].unique())))

#For a list of program codes, use get_names for each code
def get_all_names(codes):
    """
        For a list of program codes, use get_names for each code
    """
    return {code: get_names(code) for code in codes}

#Get the program short title from df_pgrm_blk give a program code
def get_prg_name(code):
    """ 
        Get the program short title from df_pgrm_blk give a program code
    """
    if code == 'No Match':
        return "No Match"
    try:
        return df_pgrm_blk[df_pgrm_blk['Program Code'] == code]['Program Short Title'].iloc[0]
    except:
        return "No Match"

#Given a unique combination of course and section number, designated as NC,
#find the program code from df_pgrm_blk
def find_codes(NC):
    """
        Given a unique combination of course and section number, designated as NC,
        find the program code from df_pgrm_blk
    """
    if (df_pgrm_blk[df_pgrm_blk['NC'] == NC]['Program Code'].empty):
        return ['No Match']
    else:
        return list(df_pgrm_blk[df_pgrm_blk['NC'] == NC]['Program Code'].unique())


#Isolate the section number from a string containing the section number as the only numeric part
def isolate_sec(sec_string):
    """Isolate the section number from a string containing the section number as the only numeric part"""
    return ''.join(char for char in sec_string if char.isdigit())

#Extract the emails of instructors with @conestogac.on.ca domains
def extract_work_email(email_string):
    """
        Extract the emails of instructors with @conestogac.on.ca domains
    """
    l_o_emails = list(filter(lambda x: '@conestogac.on.ca' in x, email_string.split(", ")))
    if (len(l_o_emails) == 0):
        return "No match"
    return l_o_emails[0]

# --------------END: Functions --------------------------------------

## --------------- START: script ------------------------------------

## Data input

# Get the folder containing the data files
root = Tk.Tk()
root.withdraw() #use to hide tkinter window

currdir = os.getcwd()
container = filedialog.askdirectory(parent=root, initialdir=currdir, title='Please select a directory')

# Assign Data Frames from files in the directory:
df_pgrm_blk = pd.read_excel(os.path.join(container, "Program Block Extract.xlsx"))
df_pgrm_blk = df_pgrm_blk.fillna("")

df_cntct_lst = pd.read_excel(os.path.join(container, "Program Contact List Extract.xlsx"))
df_cntct_lst = df_cntct_lst.fillna("")

df_pgrm_coord = pd.read_excel(os.path.join(container, "Program Coordinators.xlsx"))
df_pgrm_coord = df_pgrm_coord.fillna("")

df_crse_rep = pd.read_excel(os.path.join(container, 'Course Startup Report short list.xlsx'))
df_crse_rep = df_crse_rep.fillna("")

# Create unique keys by combining section number and course code (NC):
df_crse_rep['NC_helper'] = df_crse_rep['textbox5'].map(isolate_sec) + df_crse_rep['textbox80']    
df_pgrm_blk['NC'] = df_pgrm_blk['Section'].map(isolate_sec) + df_pgrm_blk['Course'].map(lambda x: x.split(' ')[0])

# list of all NCs:
l_o_NC = list(df_crse_rep['NC_helper'].unique())

# Dictionary of all names and emails:
huge_dict = {NC: get_all_emails(get_all_names(find_codes(NC))) for NC in list(df_crse_rep['NC_helper'].unique())}

#Create a dictionary of just program and coordinators
dict_of_coord = {}
for NC in huge_dict.keys():
    for  code in huge_dict[NC].keys():
        dict_of_coord[code] = [(name, email) for name, email in huge_dict[NC][code][1].items()]

# Update the Program Block Extract with the new information
# Helper Functions --------------------
def get_coord1_name(code):
    if code not in dict_of_coord:
        return "No Match"
    return dict_of_coord[code][0][0]

def get_coord1_email(code):
    if code not in dict_of_coord:
        return "No Match"
    return dict_of_coord[code][0][1]

def get_coord2_name(code):
    if code not in dict_of_coord:
        return "No Match"
    if len(dict_of_coord[code]) > 1:
        return dict_of_coord[code][1][0]
    else:
        return ""
    
def get_coord2_email(code):
    if code not in dict_of_coord:
        return "No Match"
    if len(dict_of_coord[code]) > 1:
        return dict_of_coord[code][1][1]
    else:
        return ""
#-----------------------    
# Update the Program Block Extract with the new information
df_pgrm_blk['Coord1'] = df_pgrm_blk['Program Code'].map(get_coord1_name)
df_pgrm_blk['Coord1 email'] = df_pgrm_blk['Program Code'].map(get_coord1_email)
df_pgrm_blk['Coord2'] = df_pgrm_blk['Program Code'].map(get_coord2_name)
df_pgrm_blk['Coord2 email'] = df_pgrm_blk['Program Code'].map(get_coord2_email)

# Save the formatted program block extract:
df_pgrm_blk.to_csv(os.path.join(container, "Formatted Program Block Extract.csv"), index = False)

#Get the unique number of coordinators for each section:
dict_unique_coord = {}

for NC in huge_dict.keys():
    dict_unique_coord[NC] = {}
    for code in huge_dict[NC].keys():
        dict_unique_coord[NC] = {**dict_unique_coord[NC], **huge_dict[NC][code][1]}

# Put in all the coordinator Names and Emails in the course report
# Helper Function:
def get_coord_names_emails(NC):
    temp_list = [(name, email) for name, email in dict_unique_coord[NC].items()]
    
    while len(temp_list) < 4:
        temp_list.append(("",""))
    
    return temp_list[0][0], temp_list[0][1], temp_list[1][0], temp_list[1][1], \
            temp_list[2][0], temp_list[2][1], temp_list[3][0], temp_list[3][1],
#--------------------
# Put in all the coordinator Names and Emails in df_crse_rep
df_crse_rep[['Coord 1 Name', 'Coord 1 Email',\
      'Coord 2 Name', 'Coord 2 Email',\
      'Coord 3 Name', 'Coord 3 Email',\
      'Coord 4 Name', 'Coord 4 Email']] = pd.DataFrame(df_crse_rep['NC_helper'].apply(get_coord_names_emails).apply(pd.Series))

# Get Instructors' work emails
df_crse_rep['Instructor Email'] = df_crse_rep['Home_Email'].map(extract_work_email)

# Get all the instructor names in the right place
# List of rows to remove
index_list = []

# Helper Function ---------------------------------
def get_instructor(NC):
    temp1 = df_crse_rep[df_crse_rep['NC_helper'] == NC]
    temp2 = temp1["textbox5.1"]
    idx = temp2.apply(lambda x: x if x else np.NaN).first_valid_index()
    if idx == None:
            return "No Match, No Match", "No Match"
    instructor =  temp2[idx]
    email = temp1["Instructor Email"][idx]
    index_list.append(idx)
    return instructor, email
#-------------------
#Get all the instructor names in the right place
df_crse_rep[['Instructor', 'Instructor Email']] = pd.DataFrame(df_crse_rep['NC_helper'].apply(get_instructor).apply(pd.Series))
df_crse_rep.drop(index_list, inplace = True)

# Use Instructor First and Instructor Last for instructor names:
df_crse_rep['Instructor First Name'] = df_crse_rep['Instructor'].apply(lambda x: x.split(", ")[1])
df_crse_rep['Instructor Last Name'] = df_crse_rep['Instructor'].apply(lambda x: x.split(", ")[0])

#Separate Section Number and Campus
df_crse_rep['Section'] = df_crse_rep['textbox5'].apply(lambda x: "Section "+ x.split(" ")[0].strip("#"))
df_crse_rep['Campus'] = df_crse_rep['textbox5'].apply(lambda x: x.strip().split(" ")[-1])

# Create Formatted Course Report:
formatted_crse_rep = df_crse_rep.iloc[:, [0, 2, 5, 6, 9, 13, 14, 15, 16, 17]]
formatted_crse_rep = pd.concat([formatted_crse_rep,\
           df_crse_rep.loc[:, ['Instructor First Name', 'Instructor Last Name',
            'Section', 'Campus',
            'Instructor Email', 'Coord 1 Name',
            'Coord 1 Email',
            'Coord 2 Name',
            'Coord 2 Email',
            'Coord 3 Name',
            'Coord 3 Email',
            'Coord 4 Name',
            'Coord 4 Email']]], axis = 1)

# Rename the columns

formatted_crse_rep[['Course Code', 'Section_End_Date', 'Number Enrolled', 'Class Type', 'Day of Week', 'Lecture Start', 'Duration']]\
    = formatted_crse_rep[['textbox80', 'textbox32', 'textbox145', 'textbox14', 'textbox46', 'textbox7', 'textbox12']]

formatted_crse_rep = formatted_crse_rep[['Course Code', 'Section',\
 'Campus', 'Section_Start_Date', 'Section_End_Date',\
 'Number Enrolled',
 'Class Type',
 'Day of Week',
 'Lecture Start',
 'Duration',
 'Instructor First Name',
 'Instructor Last Name',
 'Instructor Email',
 'Coord 1 Name',
 'Coord 1 Email',
 'Coord 2 Name',
 'Coord 2 Email',
 'Coord 3 Name',
 'Coord 3 Email',
 'Coord 4 Name',
 'Coord 4 Email']]

# Remove unnecessary info from enrollment number
formatted_crse_rep['Number Enrolled'] = formatted_crse_rep['Number Enrolled'].apply(lambda x: str(x).split()[0])

# Save final output to file
formatted_crse_rep.to_csv(os.path.join(container, "Formatted Course Report.csv"))