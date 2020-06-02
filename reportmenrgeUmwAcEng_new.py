# filname: reportmerge.py

## to merge 3 csv files useful for report writing.
## May 2018
## All files must be .csv
## names file must be one name per line
## other files must be as exported by Moodle.

from pandas import *

def import_data(): #so far only string-literals, not imput dialogues
    print('This program takes a csv file with a list of student names and "Email address"')
    print('There must be a file from moodle in this directory called acengmarks.csv')
    print('and another called umwmarks.csv')
    filename_in = input('What is the name of your names + email file: ')
    names = read_csv(filename_in)
    umw = read_csv('umwmarks.csv')
    eng = read_csv('acengmarks.csv')

    return names, umw, eng

def merge_dataframes(names, umw, eng):
    temp_merge = pandas.merge(names, eng, how='left', on='Email address')
    merged = pandas.merge(temp_merge, umw, how='left', on='Email address')
    return merged

def rename_columns(df):
    df = df.rename(columns={'Folder Term 1 (Real)': 'T1 folder'
                            , 'Tourism essay (Real)': 'T1 essay: tourism'
                            , 'Internet Democracy (Real)': 'T2 essay: internet'
                            , 'Mock re-write (Real)': 'T3 essay: obesity'
                            , 'Turnitin Assignment 2: 15IC-UUMW-A17/18-AS1 (Real)': 'UMW T1'
                            , 'Turnitin Assignment 2: 15IC-UUMW-A17/18-AS2 (Real)': 'UMW T2'
                            })
    return df

def save_as(finaldata):
    filename_out = input('What name to save data, including .csv extension: ')
    finaldata.to_csv(filename_out,  index=False)

names, umw, eng = import_data()
import pdb; pdb.set_trace()
merged = merge_dataframes(names, umw, eng)
finaldata = rename_columns(merged)
save_as(finaldata)


