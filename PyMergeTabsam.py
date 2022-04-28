# -*- coding: utf-8 -*-
# PyTabsam
# @author: sszsth, sszgrm

import json
import pandas as pd
import datetime
import shutil


# Leere Listen vorbereiten
data_files  = pd.DataFrame([],dtype=pd.StringDtype())
data_sheets = pd.DataFrame([],dtype=pd.StringDtype())

# Global variable from configuration
path_input = ""
filename_output = ""

# Function tolog
# Write logging information
def tolog(level, text):
  dateTimeObj = datetime.datetime.now()
  timestamp = dateTimeObj.strftime("%Y-%m-%d %H:%M:%S%z")
  print(level + " (" + timestamp + "): " + text)


# Function read_config
# Read the configuration file
def read_config():
  with open('config.json', 'r', encoding="utf-8") as f:
    config = json.load(f)
    list_files = []
    list_sheets = []
    global data_files, data_sheets, path_input, filename_output
    
    path_input  = config['path_input']
    filename_output = config['filename_output']
    
    for key in config:
      conf_value = config[key]
      if key == "files":
        # The files are provided as a list. Read them an add them do the Dataframe data_files
        for i in range(len(conf_value)):
          files_elem = conf_value[i]
          pk = i+1
          input_fullpath = path_input + "/" + files_elem["input_filename"]
          elem_list_files = [pk, files_elem["title"], input_fullpath, files_elem["position"]]
          list_files.append(elem_list_files)
          data_files = pd.DataFrame(list_files, columns = ['id', 'title' , 'input_path', 'position'])
      if key == "sheets":
        # The sheets to process are provided as a list. Read them an add them do the Dataframe data_sheets
        for i in range(len(conf_value)):
          sheets_elem = conf_value[i]
          pk = i+1
          elem_list_sheets = [pk, sheets_elem["code"], sheets_elem["title"], sheets_elem["column"]]
          list_sheets.append(elem_list_sheets)
          data_sheets = pd.DataFrame(list_sheets, columns = ['id', 'code', 'title', 'column'])


# Function create_tabsam 
# Create and generate the destination excel files
# If the destination files already exists, they will be overwritten
def create_tabsam():
  global filename_output

  # Create empty output file based on template
  tolog("INFO", "Writing output to: " + filename_output)
  shutil.copy('VorlageTabsam.xlsx', filename_output)

  sheet_id = 0
  for sheet_i, sheet_row in data_sheets.iterrows():
    sheet_id = sheet_row['id']
    print(sheet_id)

    for file_i, file_row in data_files.iterrows():
      file_id = file_row['id']
      print(file_id)
    

# Main progam
def main():
  tolog("INFO", "Read the configuration")
  read_config()
  
  tolog("INFO", "Loop over the input sheets and files and merge the tables")
  create_tabsam()


# Execute main of PyTabsam
if __name__ == '__main__':
  main()
