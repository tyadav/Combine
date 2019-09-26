# Importing required packages 
import pandas as pd
import os 
import csv

# Function to get all the paths for a particular file name. Checks all the directories and sub-directories 
def get_paths(folder_name,file_name):    
    folder_path = "C:/Users/TejYadav/Desktop/Merge/"+(folder_name)+"/" # folder path 
    all_paths = []    
    for r, d, f in os.walk(folder_path):
        for file in f:
            if (file_name) in file:
                all_paths.append(os.path.join(r, file))   
    return all_paths
    
# Function to generate new column names in case of multiple files with same attribute - format : filename_attributename
def new_columns(fname,cnames):
    new_cnames = []
    for i in cnames:
        new_cnames.append(fname+"_"+i)
    return new_cnames
    
# To generate the output file for a particular file name present in a folder named folder name and extract attributes defined by attribute name 
# output file format - main_dir_path + sheet_dir_path + input_file_name.csv
def generate_file(folder_name,file_names,attribute_name,input_name):    
    dframe_file = pd.DataFrame()
    output_dframe = pd.DataFrame()    
    attribute_name = attribute_name.encode('utf-8')
    attribute_name = attribute_name.strip()
    attribute_names = attribute_name.split("\n")    
        
    for fname in file_names:
        all_paths_files = get_paths(folder_name,fname)   # get all possible paths for a folder and file name        
        if len(file_names)>1:
            temp_fname = fname.replace(".csv","")
            new_attribute_names = new_columns(temp_fname,attribute_names)  # get attributes as file_name_attribute_name for multiple files with same attribute names

        for j in range(len(all_paths_files)):  
            dframe = pd.read_csv(all_paths_files[j])            
            if j == 0:
                dframe_file = dframe.loc[:,attribute_names]                
                if len(file_names)>1:
                    dframe_file.columns = new_attribute_names
            else:
                # append to same data frame for multiple files available with different dates                
                temp_df = dframe.loc[:,attribute_names]                 
                if len(file_names)>1:
                    temp_df.columns = new_attribute_names
                dframe_file = dframe_file.append(temp_df,ignore_index=True)                
        output_dframe = pd.concat([output_dframe,dframe_file],axis=1)        
    return output_dframe    
    
# Function that uses other functions to extract all the attributes and stores output data frame corresponding to each sheet in a list
def extract_cols(sheets,output_dataframes_list):    
    for sheet_name in sheets:
        print(sheet_name)        
        dframe_sheet = pd.read_excel(xls,sheetname=sheet_name)
        records,cols = (dframe_sheet.shape)
        dframe_for_each_sheet = pd.DataFrame()
        result_df = pd.DataFrame()
        for i in range(0,records):
            folder_name = str(dframe_sheet.loc[i,'Folder Name'])
            file_name = str(dframe_sheet.loc[i,'File Name'])
            file_names = file_name.split("\n")
            attribute_name = (dframe_sheet.loc[i,'Column Name'])
            input_name = str(dframe_sheet.loc[i,'Inputs'])
            #if isinstance(attribute_name, unicode) and folder_name != ' ' and folder_name != 'nan' and file_name != ' ' and file_name != 'nan':
            if isinstance(attribute_name, str) and folder_name != ' ' and folder_name != 'nan' and file_name != ' ' and file_name != 'nan':
                byte_attribute_name = bytes(attribute_name,'utf-8')                
                result_df = generate_file(folder_name,file_names,attribute_name,input_name) # returns for each file name in a sheet
                dframe_for_each_sheet = pd.concat([dframe_for_each_sheet,result_df],axis=1) # concat acc to each sheet
        final_records,final_cols=(dframe_for_each_sheet.shape)
        output_dataframes_list.append(dframe_for_each_sheet)  # appending to output list
        print("Number of attributes in the sheet : "+str(final_cols))
    
# To read input xlsx file - input workbook - Extract file.xlsx
input_file_path = 'C:/Users/TejYadav/Desktop/Merge/WIP_V2.1.xlsx'  # Path to input file 
xls = pd.ExcelFile(input_file_path)

# Getting names of all the worksheets from input workbook 
all_sheet_names = xls.sheet_names
print(all_sheet_names)

# Calling function for considering all the sheets in the input workbook
output_dataframes_list = []    # list of data frames according to number of sheets 
extract_cols(all_sheet_names,output_dataframes_list) 

# Writing the data corresponding to each sheet in xlsx file named output.xlsx - gets saved on desktop. 
print("Writing to output file !")
from openpyxl import load_workbook

output_file_path = 'C:/Users/TejYadav/Desktop/output.xlsx'  # output file path 
writer = pd.ExcelWriter(output_file_path, engine = 'openpyxl')

for i in range(len(all_sheet_names)):
    #print(output_dataframes_list[i].columns)
    output_dataframes_list[i].to_excel(writer,sheet_name=all_sheet_names[i],encoding='utf-8',index = False)

writer.save()
writer.close()
print("File saved successfully !")

#Writing to output file !
#File saved successfully !
