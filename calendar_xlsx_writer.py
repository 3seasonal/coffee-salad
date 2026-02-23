#!/usr/bin/env python
# -*- coding: utf-8 -*-

import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
import datetime
from typing import Dict, List, Any, Tuple, Optional, Union
import os
import logging


"""Calendar xlsx writer module
    Coffee-Salad calendar package

    This module provides cordination functions for calendar creation.
    It maintains creates and or updates calendars using xlsx files.
    
    Author:     dillonj
    Created:    2025-08-29
    Description: class for writing xlsx files .

"""

    
class CalendarXlsxCreatorError():
    def __init__(self, raised_error_message: str):
        super().__init__(
            f"Exception in the XLSX Creator: {raised_error_message}"
        )
        

class calendarXlsxCreator:
    """
    Creates a XLSX file from scratch
    See calendarXLSXUpdator for updating an existing XLSX file.
    """
    

    
    def __init__(self, config: dict, logger=None, ts=datetime.datetime.now().strftime('%Y%m%d_%H%M%S')):
        """Initialize the xlsx calenadarcreator with the path to the Excel file."""
        
        #set up logging
        self.log = logger or logging.getLogger(__name__)
        
        # set output path
        self.script_dir = os.path.dirname(os.path.abspath(__file__))

        #get general unitlity class:
        dt = datetime.now().strftime('%Y%m%d_%H%M')
        
        # Store configuration for use:
        if not config:
            self.log.error("No configuration provided to calendarXlsxCreator.")
            raise CalendarXlsxCreatorError("Oh no. The XLSX writer was not passed a configuration. :(")
        else:
            self.config = config # store the config for later use, but it is not used directly in this class. It is passed to the calendarXLSXUpdator for processing.
            # contains sub-dictionaries: calendar, columns, events, styles.
       
                   
        # Handle the creation of the calendar XLSX file. If it does not exist,
        self.xlsx_name = config.get("calendar", {}).get("xlsx_name")
        
        # check if it ends with .xlsx
        if not self.xlsx_name.endswith(".xlsx"):
            self.xlsx_name = self.xlsx_name + ".xlsx"
        self.xlsx_name = self.xlsx_name.replace(".xlsx", "{dt}.xlsx") # add time stamp to the file name 
    
        # calculate path to the xlsx file:
        self.xlsx_path = os.path.join(self.script_dir, self.xlsx_name)
        self.log.info(f"XLSX file path set to: {self.xlsx_path}")    
        
        # create new xlsx file using openpyxl (check if it exists first, if it does, delete it and create a new one)
        if os.path.exists(self.xlsx_path):
            self.log.warning(f"XLSX file already exists at {self.xlsx_path}. It will be overwritten.")
            os.remove(self.xlsx_path)
        self.workbook = openpyxl.Workbook()
        self.log.info(f"New XLSX workbook created at {self.xlsx_path}.")
        
        
        
        
    def save(self):
        """Save the workbook to the specified path."""
        try:
            self.workbook.save(self.xlsx_path)
            self.log.info(f"Workbook saved successfully at {self.xlsx_path}.")
        except Exception as e:
            self.log.error(f"Failed to save workbook: {e}")
            raise CalendarXlsxCreatorError(f"Failed to save workbook: {e}")
        
## algorithm - to be implemetned later:

'''
create worksheet

save


functions:
- create worksheet
- save workbook

- update cell style
- add style to workbook

- event_is_multiday

---
create matrix.
    create trackers:
        catagory list - index of each catagory releative to the date row
        date index - given a date, return the row and col index
        days of the week 2 column - given a day of the week which column
        column 2 day of the week - given a column, return the day of the week index.
        day of the weeek - iso index
        list of styles created
        
    create styles
        and add to style created list
        
    populate numbers and dates in calendar.
        and colour the cells as they are created.
    
    
iterate events and populate the cells with the event data.
    colour approparietately
    
    event matrix list - a dictionary of of key dates, 
            each with a list of catagories.
                each catagory with a list of events. (that are the key wihtin the config.events)
                    each event is a touple of (event name, #day, of #days)
    (in this way can events for colouring purposes)
    
    
Add event

    



'''

def create_worksheet(self, sheet_name: str):
    
    pass
    # create a new worksheet with the given name
    
    # initialise the worksheet
    
    # set the current worksheet to the new worksheet (self.worksheet = self.workbook[sheet_name])
    
    
def is_workbook_open(file_path):
    """
    Check if a specific Excel workbook is currently open.

    Args:
        file_path (str): The full path to the Excel workbook file.

    Returns:
        bool: True if the workbook is open, False otherwise.

    Raises:
        psutil.NoSuchProcess: If the process no longer exists.
        psutil.AccessDenied: If the process cannot be accessed.
    """
    for proc in psutil.process_iter(['pid', 'name', 'open_files']):
        try:
            if proc.info['name'] == 'EXCEL.EXE':
                for file in proc.info['open_files'] or []:
                    if file.path == file_path:
                        return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return False

def close_workbook(file_path):
    """
    Closes the Excel workbook specified by the file path if it is currently open.

    This function iterates through all running processes to find instances of 'EXCEL.EXE'.
    If the specified file is found to be open by any of these processes, the process is terminated.

    Args:
        file_path (str): The full path to the Excel workbook to be closed.

    Raises:
        psutil.NoSuchProcess: If the process no longer exists.
        psutil.AccessDenied: If the process cannot be accessed due to permission issues.

    Returns:
        None
    """
    for proc in psutil.process_iter(['pid', 'name', 'open_files']):
        try:
            if proc.info['name'] == 'EXCEL.EXE':
                for file in proc.info['open_files'] or []:
                    if file.path == file_path:
                        proc.terminate()
                        proc.wait()
                        return
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue