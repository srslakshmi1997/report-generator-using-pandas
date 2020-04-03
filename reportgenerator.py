# -*- coding: utf-8 -*-
"""
Created on Mon May 13 17:45:12 2019

@author: sradhakr
"""

import os
import glob
import pandas as pd
import sys


def func_convert_to_csv(folder_path,result_glob) :
    ##########print(os.getcwd(), " before changing the path to the python script location") ##########
    os.chdir(os.path.abspath(os.path.dirname(sys.argv[0])))
    ##########print(os.getcwd(), " after changing the path to the python script location") ##########
    for filename in result_glob:
        os.system('''XlsToCsv.vbs "'''+str(folder_path)+'\\'+str(filename)+'''" "'''+(str(folder_path)+'\\'+str(filename)).split('.')[0]+'''.csv"''')
        os.remove(str(folder_path)+'\\'+str(filename))
    print("Conversion completed!")
    

def get_input():
    search_column = input('''Enter : \'1\' to generate report based on \'ID\' \n
        \'2\' to generate report based on \'State\' \n
        \'3\' to generate report based on a key term in \'Description\'\n''')
    search_column_dict = {1 : 'ID' , 2 : 'State', 3 : 'Description'}
    if int(search_column) == 1 :
        search_term = input("Please enter the ID : ")
    elif int(search_column) == 2 :
        search_term = input("Please enter the State : (For Example you can enter Design,Development,Testing) : ")
    else :
         search_term = input("Please enter the Search Term to perform filter in the Description : ")
    search_column_final = search_column_dict[int(search_column)]
    return (search_column_final,search_term)
    
    
def search_term_function(final_entries_dict,search_term,open_file_df,search_column):
    #Find total number of hours on that particular keyword
    total_hours = 0
    date_vals = 0
    
    if search_column == 'Description' :
        sent_list = {sentence for sentence in open_file_df[search_column].values.tolist() if search_term.lower() in sentence.lower()}
        total_hours = sum([sum(open_file_df[open_file_df[search_column] == notes]['Hours'].values) for notes in sent_list])
    elif search_column == 'State' : 
        sent_list = open_file_df[open_file_df[search_column].str.lower() == search_term.lower()].values.tolist()
        total_hours = sum([ind[5] for ind in sent_list])
        date_vals = [ind[4] for ind in sent_list]
    else : 
        sent_list = open_file_df[open_file_df[search_column].str.lower() == search_term.lower()].values.tolist()
        total_hours = sum([ind[5] for ind in sent_list])
    
    fn = file_name.split('.')
    
    #Find the name of the employee
    name_of_the_employee = ' '.join(fn[0].split()[2:4])
    
    if name_of_the_employee not in list(final_entries_dict.keys()):
        final_entries_dict[name_of_the_employee] = {}
    
    #Find the month in which the calculation was made
    indextofindmonth = int(len(open_file_df['Date'])/2)
    if '-' in open_file_df['Date'][indextofindmonth]:
        month = open_file_df['Date'][indextofindmonth].split('-')[0]
    elif '/' in open_file_df['Date'][indextofindmonth]:
        month = open_file_df['Date'][indextofindmonth].split('/')[0]
    
    #identify the Quater with the month
    if (int(month) <= 3):
        quarter = 'Q1'
    elif (int(month) > 3) and (int(month) <=6):
        quarter = 'Q2'
    elif (int(month) > 7) and (int(month) <= 9) :
        quarter = 'Q3'
    else :
        quarter = 'Q4'
    
    return(final_entries_dict,name_of_the_employee,total_hours,month,quarter,date_vals)
    
logf = open("download.log", "a")

try:
    print("Hi, Welcome to report generator!")
    terminate_flag = ''
    while terminate_flag != '2' :
        folder_path = input("Please enter the folder path where the data is stored : " )
        ##########print(os.getcwd(),"before changing to user entered directory") ########
        os.chdir(folder_path)
        ##########print(os.getcwd(), "after the change to the user entered folder") ######
        extension = 'xlsx'
        result_glob = glob.glob('*.{}'.format(extension))
        if result_glob :
            print("Converting the .xlsx files to .csv....")
            func_convert_to_csv(folder_path,result_glob)
        
        # Change the directory to the folder location
        ##########print(os.getcwd(), " all csv or return from function, change back to user entered directory") ##########
        os.chdir(folder_path)
        ##########print(os.getcwd(), "after the change to the user entered folder to work on the csv files") #########
    
        # Get all the files in the located folder as a list
        files_list = glob.glob('*.csv')
    
        search_column,search_term = get_input()
    
        final_entries_dict = {}
        data_dict = {}
        for file_name in files_list :
            open_file_df = pd.read_csv(file_name,encoding = "ISO-8859-1")
            final_entries_dict,name_of_the_employee,total_hours,month,quarter,date_vals = search_term_function(final_entries_dict,search_term,open_file_df,search_column)
            #data_dict[quarter] = total_hours
            #final_entries_dict[name_of_the_employee] = data_dict
            #final_entries_dict.update({name_of_the_employee : {quarter : total_hours }})
            final_entries_dict[name_of_the_employee][quarter] = total_hours
            if  date_vals != 0:
                final_entries_dict[name_of_the_employee]['Date '+str(quarter)] = date_vals
            
            #print(final_entries_dict)
        write_dataframe = pd.DataFrame.from_dict(final_entries_dict,orient='index')
        #print(write_dataframe)
        output_filename = input("Please provide the output file name : ")
        output_directory = input("Please provide the output file directory : ")
        ##########print(os.getcwd(), "before the change to the user entered folder to print the csv files") #########
        os.chdir(output_directory)
        ##########print(os.getcwd(), "after the change to the user entered folder to print the csv files") #########
        write_dataframe.to_excel(output_filename+".xlsx")
        print("The file "+output_filename+".xlsx"+" has been generated successfully under the folder path "+output_directory )
        terminate_flag = input("Do you want to generate another report ? Enter '1' for 'Yes', '2' for 'No' ")

except Exception as e:
     logf.write("Error: "+str(e))
     print("Something went wrong!")
    
    
    
    
    


