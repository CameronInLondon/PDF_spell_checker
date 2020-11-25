#!/usr/bin/env python3

# TODO - refactor code

import glob
import fitz                   # this is PyMuPDF
from docx import Document     # python.docx library
import win32com.client
import os
import pandas as pd
from timeit import default_timer as timer


### inputs and outputs
location_of_pdf_files = r"C:\file_path\*.pdf"
output_text_files = r"C:\file_path\*.txt"
output_csv_file = r'C:\file_path\spelling_data_output.csv'
input_data_set = r"C:\file_path\input_date.xlsx"

#  VERSION 2 of this program
print('--------------------- Welcome to PDF spellerChecker -----------------------------')
print('--------------------- Developed by the CPS data team ----------------------------')
log_number = (input('Please type the name of the column with the log numbers:') or 'log_number')
file_path = (input('please type the name of the column with the full file paths:') or 'file_path')
sector = (input('please type the name of the column with the sector detail:') or 'Sector')
syllabus_no = (input('please type the name of the column with the syllabus number:') or 'Syllabus No')
print('------ This program normally takes about 10 seconds per PDF per 10 page PDF -----')
print('-------------------------- Now sit back and wait --------------------------------')

def iterate_file_path_df():
    """
    Reads excel file into df and deletes all unwanted columns

            Parameters:
                    input_date_set (global):

            Returns:
                    call main()
    """
    df = pd.read_excel(input_data_set)
    col_list = [file_path, log_number, sector, syllabus_no]
    # filter out all columns bar required ones
    df = df[col_list]
    main(df)


def main(df):
    start = timer()
    df3 = pd.DataFrame(columns=['notes','log_number', 'sector', 'syllabus', 'spelling_errors', 'grammar_errors' ,'file_path_full'])
    df2 = pd.DataFrame(columns=['notes','log_number', 'sector', 'syllabus', 'spelling_errors', 'grammar_errors', 'file_path_full' ])
    for index, row in df.iterrows():
        row_of_cols = [row[log_number], row[sector], row[syllabus_no], row[file_path]]
        log = int(row[log_number])
        if log == '':
            end = timer()
            print('Finished!!! Program took: ', round(end - start, 0), 'seconds to run') # Time in seconds
            exit()
        print(row[log_number])
        print(row[file_path] )  
        if row[file_path] != 'end':
            if row[file_path].endswith(".pdf"):
                if os.path.isfile(row[file_path]):
                    text = get_text_from_pdf(row[file_path])
                    output = create_doc_and_write(text)
                    pass_error_lists = spelling_checker(output)
                    pass_error_lists = check_error_list(pass_error_lists)
                    pass_error_lists = convert_lists_to_strings(pass_error_lists)
                    df2 = save_to_df(pass_error_lists, *row_of_cols)
                    df3 = df3.append(df2, ignore_index=True)
                    print('log number: ', log,  "---break 7 - save to file---")
                    save_to_file(df3)
                else:
                    loc_append_and_save(df2, df3, *row_of_cols) 
                    save_to_file(df3)
            else:
                print("----get_files_function----")
                latest_file = get_files(row[file_path], row[log_number])
                if 'no log number provided' or 'no_path' in latest_file:
                    loc_append_and_save(df2, df3, *row_of_cols)
                    save_to_file(df3)
                else:
                    text = get_text_from_pdf(row[file_path])
                    output = create_doc_and_write(text)
                    pass_error_lists = spelling_checker(output)
                    pass_error_lists = check_error_list(pass_error_lists)
                    pass_error_lists = convert_lists_to_strings(pass_error_lists)
                    df2 = save_to_df(pass_error_lists, *row_of_cols)
                    print(df2)
                    df3 = df3.append(df2, ignore_index=True)
                    print('log number: ', log,  '---break 7 - save to file---')
                    save_to_file(df3)
        else:
            print('-----------all done------------')
            end = timer()
            print('Program took: ', round(end - start, 0), 'seconds to run..') # Time in seconds
            exit()


def save_to_file(df3):
    """
    Function to save df to csv

            Parameters:
                    df3 (df):
                    encoding (str): optional -- Note there is a pandas bug with to_csv passing encoding      
                                    TypeError: "delimiter" must be a 1-character string

            Returns:
                    n/a
    """
    df3.to_csv(output_csv_file, encoding='utf-8-sig')


def loc_append_and_save(df2, df3, log, sector, syllabus, current_file_path):
    if log < 1:
        df2.loc[0] = ['no log number provided', log, sector, syllabus,'n/a','n/a', current_file_path]
    else:
        df2.loc[0] = ['file renamed or moved', log, sector, syllabus ,'n/a','n/a', current_file_path]
    df3 = df3.append(df2, ignore_index=True) 
    print('log number: ', log , '---save to file---')
    save_to_file(df3)


def get_files(path, log):
    """
    Function to use given folder location iterate through folder(s) grabing .PDF files with same 'log number' and output to a list.

            Parameters:
                    path (str):
                    log (str):

            Returns:
                    latest_file (list): 
    """
    if log < 1:
        latest_file = 'no log number provided'
        return latest_file
    list_of_file_paths = []
    half_file_path = path + r'**\*' + str(log) +'*.pdf'
    print('------ searching drive for file could take some time ------')
    list_of_file_paths = glob.glob(half_file_path, recursive=True)
    if list_of_file_paths:
        latest_file = max(list_of_file_paths, key=os.path.getmtime)
    else:
        latest_file = 'no_path'
    return latest_file


def get_text_from_pdf(next_pdfs_path):
    """
    Returns (str)

            Parameters:
                    next_pdfs_path (str): file path use r"" around path.

            Returns:
                    text (str): string of text
    """
    if os.path.isfile(next_pdfs_path):
        try:
            text = ''
            # using PyMuPDF library to get text.
            with fitz.open(next_pdfs_path) as doc:
                for page in doc:
                    text+= page.getText()
            return text
        except (RuntimeError, IOError):
            pass
    pass


def create_doc_and_write(text):
    """
    Returns a tuple with two win32com objects

            Parameters:
                    text (str): this is the text that needs to be spell checked. One LONG str.

            Returns:
                    output (tuple): two items (text document, word.applicaiton)
    """
    # create a new instance of Word, using Early Binding
    app = win32com.client.gencache.EnsureDispatch('Word.Application')
    print(app)
    # create a new word document
    doc = app.Documents.Add()
    # add content
    rng = doc.Range(0,0)
    rng.InsertAfter(text)
    output = (doc, app)
    return output


def spelling_checker(input_tuple):
    """
    Returns list of list with two lists

            Parameters:
                    input_tuple (tuple): tuple with two lists win32com objects

            Returns:
                    all_errors (list): contains gramar_list and spelling_list
    """
    (doc, app) = input_tuple
    # VBR
    grammar =  "Grammar: %d" % (doc.GrammaticalErrors.Count,)
    spelling =  "Spelling: %d" % (doc.SpellingErrors.Count,)
    # working with errors in VBR
    wdDoNotSaveChanges = 0
    spelling_list = []
    grammar_list = []
    i = 1
    for i in range(1, doc.SpellingErrors.Count+1):
        try:
            spelling_list.append(doc.SpellingErrors.Item(i).Text)
        except:
            print('no more spelling errors')
        try:
            grammar_list.append(doc.GrammaticalErrors.Item(i).Text)
        except:
            print('no grammer errors')
    print ('grammar_list:')
    print (*grammar_list)
    print ('spelling_list:')
    print (*spelling_list) # using * to print the list
    app.Quit(wdDoNotSaveChanges)
    all_errors = [grammar_list, spelling_list]
    return all_errors


def convert_lists_to_strings(all_errors):
    """
    convert the lists ['word1', 'word2', 'word3'] to string with line spaces 'word1\nword2\nword3\n'
    """
    spelling_string = "\n".join(all_errors[1])
    grammar_string = "\n".join(all_errors[0])
    all_errors = [grammar_string, spelling_string]
    return all_errors


def check_error_list(all_errors):
    """
    Function to check if words are protected and delete from error_list if they are.
    Example: MS word thinks (i) is an error but it is not.
            Parameters:
                    all_errors (list): list of list with two list [[grammars_list],[spellings_list]]
            
            Inputs:
                    protected_word (str/list): this is the location to add more protected words

            Returns:
                    all_errors (list): list of list with two list [[grammars_list],[spellings_list]]
    """
    [grammar_list, spelling_list] = all_errors
    protected_word = 'i'
    # check if protected word is in list
    if protected_word in spelling_list:
        # iterate list
        for i in spelling_list[:]:       ### [:] makes a copy of the list and saves the need for a line like this:   copy_spelling_list = spelling_list
            # check if iterator == protected word if yes remove
            if i == protected_word:
                # remove word from original spelling_list
                spelling_list.remove(i)
    all_errors = [grammar_list, spelling_list]
    print('>>>> output after check_error_list function: ', *spelling_list)
    return all_errors


def save_to_df(all_errors, log, sector, syllabus, current_file_path):
    """"
    Returns dataframe with two lists added.

            Parameters:
                    all_errors (tuple): tuple with two lists ([grammars_list],[spellings_list])

            Returns:
                    df (df): dataframe
    """
    [grammars_list, spellings_list] = all_errors    #, filepathlist2
    df = pd.DataFrame(columns=['notes','log_number', 'sector', 'syllabus', 'spelling_errors', 'grammar_errors', 'file_path_full'])
    df.loc[0] = ['n/a', log, sector, syllabus, spellings_list, grammars_list, current_file_path]
    print('----save_to_df_function---')
    return df



if __name__ == "__main__":
    iterate_file_path_df()


