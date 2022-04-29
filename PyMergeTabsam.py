# -*- coding: utf-8 -*-
# PyTabsam
# @author: sszsth, sszgrm

import json
import pandas as pd
import datetime
import shutil
import openpyxl
from openpyxl.styles import Font
from copy import copy


# Leere Listen vorbereiten
data_files  = pd.DataFrame([],dtype=pd.StringDtype())
data_sheets = pd.DataFrame([],dtype=pd.StringDtype())

# Global variable from configuration
path_input = ""
filename_output = ""
row_start = 9

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
      
      if file_row['position']=="1":
        print("Tabelle vorbereiten")
        prepare_table(file_row, sheet_row)
      else:
        print("ergänzen")

# Create a new table with header and index column
def prepare_table(file_row, sheet_row):
  global row_start, filename_output

  # Opening the destination xlsx and create the new worksheet
  dest_wb = openpyxl.load_workbook(filename_output)
  dest_ws = dest_wb["T_1"]
  
  table_title = sheet_row['code'] + " " + sheet_row['title']
  dest_ws.cell(row=row_start, column=1).value = table_title
  dest_ws.cell(row=row_start, column=1).font = Font(name='Arial', size=8)
  row_start += 2

  # Opening the source xlsx
  source_xlsx = file_row['input_path']
  print(source_xlsx)
  source_wb = openpyxl.load_workbook(source_xlsx)
  worksheet = sheet_row['code']
  
  # Check if worksheet exists
  if worksheet not in source_wb.sheetnames:
    tolog("ERROR", "Worksheet " + worksheet + " does not exist in " + source_xlsx)
  else:
    source_ws = source_wb[worksheet]
    
    row_source = 10
    # Copy the top left cell of the primary table
    topleft_cell = source_ws.cell(row=row_source, column=1)
    dest_ws.cell(row=row_start, column=1).value = topleft_cell.value
    dest_ws.cell(row=row_start, column=1).font  = copy(topleft_cell.font)
    
    # Find the single relevant column
    column_pos = 2
    while True:
      scan_cell = source_ws.cell(row=row_source, column=column_pos)
      # Convert numeric header cells (eg. years) to character
      if isinstance(scan_cell.value, int):
        scan_cell_value = str(scan_cell.value)
      else:
        scan_cell_value = scan_cell.value
      if scan_cell_value == sheet_row['column']:
        break
      column_pos += 1
      if column_pos > 20:
        column_pos = 0
        tolog("ERROR", "Column " + sheet_row['column'] + " does not exist in worksheet " + worksheet + " of " + source_xlsx)
        break
    
    # Copy the the header column and the single relevant column
    if column_pos > 0:
      # Set a title as the column header for the single relevant column
      dest_ws.cell(row=row_start, column=2).value = file_row['title']
      # Apply the font settings of the top left column
      dest_ws.cell(row=row_start, column=2).font  = copy(topleft_cell.font)
      
      while True:
        row_source += 1
        row_start  +=1
        head_cell = source_ws.cell(row=row_source, column=1)
        if head_cell.value is None:
          # End of data rows
          break
        if row_source > 100:
          tolog("WARNING", "No end of header column found in worksheet " + worksheet + " of " + source_xlsx)
          break
        dest_ws.cell(row=row_start, column=1).value = head_cell.value
        dest_ws.cell(row=row_start, column=1).font  = copy(head_cell.font)
        # copy alignment of header cell
        dest_ws.cell(row=row_start, column=1).alignment = copy(head_cell.alignment)
        data_cell = source_ws.cell(row=row_source, column=column_pos)
        dest_ws.cell(row=row_start, column=2).value = data_cell.value
        dest_ws.cell(row=row_start, column=2).font  = copy(data_cell.font)
        # copy alignment of data cell
        dest_ws.cell(row=row_start, column=2).alignment = copy(data_cell.alignment)
        # copy number_format of data cell
        dest_ws.cell(row=row_start, column=2).number_format = copy(data_cell.number_format)
    
    row_start += 4

  dest_wb.save(filename_output)


# Main progam
def main():
  tolog("INFO", "Read the configuration")
  read_config()
  
  tolog("INFO", "Loop over the input sheets and files and merge the tables")
  create_tabsam()


# Execute main of PyTabsam
if __name__ == '__main__':
  main()
