#! /usr/bin/env python3


"""
This is a program to merge csv files into a single Excel spreadsheet.
It requires 2 input file types:
  - a csv file downloaded from moodle gradebook for each of
    academic English, UMW and (optionally) QM.
    - a csv list of student emails & names. The format must be:
        Email address,forename,surname
        654732@soas.ac.uk,John,Smith
        etc.

Dataframes are merged on the 'Email address' column.

The final Excel .xlxs file will be used to mailmerge report forms.
"""

import os
from pandas import *


def import_data():
    all_csv = os.listdir()
    # Write out instructions.
    print('This program takes a csv file with a list of student names and "Email address"')
    print('There must be a file from moodle in this directory called acengmarks.csv')
    print('and others called umwmarks.csv and qmmarks.csv')
    print()
    print('List of possible file names:')
    for csv_name in all_csv:
        if csv_name.endswith('csv'):
            print(csv_name)
    print()

    # Choose list of student names.
    filename_in = input('What is the name of your <names + email> file: ')
    names = read_csv(filename_in)
    names['Email address'].str.strip() #remove whitespace on key column

    # Load QM marks if file exists.
    if 'qmmarks.csv' in all_csv:
        qm = pandas.read_csv('qmmarks.csv')
        qm['Email address'].str.strip()
        # drop unwanted columns
#        qm = qm.drop(['First name', 'Surname', 'ID number', 'Institution', 'Department',], axis=1)
    else: qm = names  # Placeholder with no new data.

    # Load UMW marks
    umw = read_csv('umwmarks.csv')
    umw['Email address'].str.strip()
#    umw = umw.drop(['First name', 'Surname', 'ID number', 'Institution', 'Department'
#                    , 'Course total (Real)', 'Last downloaded from this course'
#                    , 'Course total (Real)', 'Last downloaded from this course'], axis=1)

    # Load Academic English marks
    eng = read_csv('acengmarks.csv')
    eng['Email address'].str.strip()
#    eng = eng.drop(['First name', 'Surname', 'ID number', 'Institution', 'Department',
#        'Assignment: Essay Plan for Veil (Real)',
#       'Assignment: Self Evaluation homework (Real)',
#        'Assignment: Term 1 Evaluation (Real)',
#       'Turnitin Assignment 2: University Access (Real)','Category total (Real)',
#       'Assignment: Rewriten Paragraph from Timed Writing (faith schools) (Real)',
#       'Turnitin Assignment 2: Tourism in Developing Countries (Real)',
#       'Workshop: DRAFT Internet Democracy Essay (submission) (Real)',
#       'Workshop: DRAFT Internet Democracy Essay (assessment) (Real)',
#       'Turnitin Assignment 2: The Internet and Democracy (Real)',
#       'Turnitin Assignment 2: Mock Rewrite: Obesity (Real)',
#       'Course total (Real)', 'Last downloaded from this course'], axis=1)
    return names, umw, eng, qm

def merge_dataframes(names, umw, eng, qm):
    temp_merge = pandas.merge(names, eng, how='left', on='Email address')
    temp_merge2 = pandas.merge(temp_merge, umw, how='left', on='Email address')
    merged = pandas.merge(temp_merge2, qm, how='left', on='Email address')
    return merged

def rename_columns(df):
    df = df.rename(columns={'Folder Term 1 (Real)': 'T1 folder'
                            , 'Tourism essay (Real)': 'T1 essay: tourism'
                            , 'Internet Democracy (Real)': 'T2 essay: internet'
                            , 'Mock re-write (Real)': 'T3 essay: obesity'
                            , 'Turnitin Assignment 2: 15IC-UUMW-A17/18-AS1 (Real)': 'UMW T1'
                            , 'Turnitin Assignment 2: 15IC-UUMW-A17/18-AS2 (Real)': 'UMW T2'
                            , 'QM Term 2 Examination Results (Real)': 'QM Term2'
                            ,  'QM Term 1 Examination Results (Real)': 'QM Term1'
                            })
    return df

def save_as(finaldata):
    filename_out = input('Choose a name for the new spreadsheet: ')
    if filename_out.endswith('.xlsx'):
        pass
    else: filename_out = filename_out + '.xlsx'
    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pandas.ExcelWriter(filename_out, engine='xlsxwriter')
    # Convert the dataframe to an XlsxWriter Excel object.
    finaldata.to_excel(writer, sheet_name='Sheet1')
    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

def main():
    names, umw, eng, qm = import_data()
    merged = merge_dataframes(names, umw, eng, qm)
    finaldata = rename_columns(merged)
    save_as(finaldata)

if __name__ == '__main__':
     main()
