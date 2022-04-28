# -*- coding: utf-8 -*-
# PyTabsam
# @author: sszsth, sszgrm

import json
import pandas as pd
import datetime


# Leere Listen vorbereiten
data_files  = pd.DataFrame([],dtype=pd.StringDtype())


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
    global data_files, path_input, path_output
    
    path_input  = config['path_input']
    filename_output = config['filename_output']
    
    for key in config:
      conf_value = config[key]
      if key == "files":
        # The collections are provided as a list. Read them an add them do the Dataframe data_coll
        for i in range(len(conf_value)):
          files_elem = conf_value[i]
          pk = i+1
          input_fullpath = path_input + "/" + files_elem["input_filename"]
          elem_list_files = [pk, files_elem["title"], input_fullpath, files_elem["position"]]
          list_files.append(elem_list_files)
          data_files = pd.DataFrame(list_files, columns = ['id', 'title' , 'input_path', 'position'])

# Main progam
def main():
  tolog("INFO", "Read the configuration")
  read_config()


# Execute main of PyTabsam
if __name__ == '__main__':
  main()
