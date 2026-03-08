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
        '''
        Simple error class for handling errors in the calendarXlsxCreator class. It takes a string message as input and formats it for display.
        
        Args:
            raised_error_message (str): The error message to be displayed when the exception is raised.
        Returns:
            None
        Raises:
            None
        '''
        super().__init__(
            f"Exception in the XLSX Creator: {raised_error_message}"
        )
        

class calendarXlsxCreator:
    """
    Creates a XLSX file from scratch
    See calendarXLSXUpdator for updating an existing XLSX file.
    """
    

    
    def __init__(self, config: dict, logger=None, ts=datetime.datetime.now().strftime('%Y%m%d_%H%M%S')):
        """Initialize the xlsx calenadarcreator with the path to the Excel file.
        
        Args:
            config (dict): The configuration dictionary containing calendar settings.
            logger (logging.Logger, optional): A logger instance for logging messages. Defaults to None.
            ts (str, optional): A timestamp string to be used in the file name. Defaults to the current date and time in 'YYYYMMDD_HHMMSS' format.
            
        Returns:
            None
            
        Raises:
            CalendarXlsxCreatorError: If there is an error during initialization, such as missing configuration.
        
        """
        
        #set up logging
        self.log = logger or logging.getLogger(__name__)
        
        # set output path
        self.script_dir = os.path.dirname(os.path.abspath(__file__))

        #get general unitlity class:
        dt = datetime.now().strftime('%Y%m%d_%H%M%S')
        
        # Store configuration for use:
        if not config:
            self.log.error("No configuration provided to calendarXlsxCreator.")
            raise CalendarXlsxCreatorError("Oh no. The XLSX writer was not passed a configuration. :(")
        else:
            self.config = config # store the config for later use, but it is not used directly in this class. It is passed to the calendarXLSXUpdator for processing.
            # contains sub-dictionaries: calendar, columns, events, styles.
                   
        # Handle the creation of the calendar XLSX file. If it does not exist,
        self.xlsx_name = config.get("calendar", {}).get("xlsx_name")
        if not self.xlsx_name.endswith(".xlsx"):
            self.xlsx_name = self.xlsx_name + ".xlsx"
        self.xlsx_name = self.xlsx_name.replace(".xlsx", f"_{dt}.xlsx") # add time stamp to the file name 
    
        # calculate path to the xlsx file:
        self.xlsx_path = os.path.join(self.script_dir, self.xlsx_name)
        self.log.info(f"XLSX file path set to: {self.xlsx_path}")    
        
                # initialise configuration handlers:
        self.events = config['events'] # the events configuration, used for populating the calendar with events
        self.category_list = list(self.events.keys()) # list of categories in the calendar, used for indexing the columns 
        self.calender_config = config['calendar']

        # create new xlsx file using openpyxl (check if it exists first, if it does, delete it and create a new one)
        if os.path.exists(self.xlsx_path):
            self.log.warning(f"XLSX file already exists at {self.xlsx_path}. It will be overwritten.")
            os.remove(self.xlsx_path)
        self.workbook = openpyxl.Workbook()
        self.log.info(f"New XLSX workbook created at {self.xlsx_path}.")
        self.worksheet = self.workbook.active
        self.worksheet.title = (self.calender_config['worksheet_name'])
        self.cell_date={} # a dictionary to store the cell references for each date

        # get day stats
        start_date = datetime.strptime(self.calender_config['start_date'], "%Y-%m-%d").date()
        end_date = datetime.strptime(self.calender_config['end_date'], "%Y-%m-%d").date()
        underflow_delta = datetime.timedelta(weeks = self.calender_config['underflow_weeks'])
        overflow_delta = datetime.timedelta(weeks = self.calender_config['overflow_weeks'])
        self.first_date = start_date - underflow_delta
        self.last_date = end_date + overflow_delta
        self.column_config = self.config['columns']
        dow_cols = self.column_config['days_of_week_columns']
        self.column_list = self.column_config['column_order']
        self.first_dow_column = self.column_list.index(dow_cols[0]) # get the column index of the first day of week column 

        # confirm the start date is the iso day of the week defined by the config:
        if start_date.isoweekday() != self.calender_config['week_stats_on']:
            self.log.critical(f"Start date {start_date} does not match the expected day of the week defined in the configuration (week_stats_on: {self.calender_config['week_stats_on']}). This may lead to misalignment of dates in the calendar.")
            raise CalendarXlsxCreatorError(f"Start date {start_date} does not match the expected day of the week defined in the configuration (week_stats_on: {self.calender_config['week_stats_on']}). Please check your configuration and try again.")

        # confirm the start_column and start_row are valid (+ve ints and not zero):
        if not isinstance(self.calender_config['start_column'], int) or self.calender_config['start_column'] <= 0:
            self.log.critical(f"Invalid start_column value: {self.calender_config['start_column']}. Please check your configuration and try again.")
            raise CalendarXlsxCreatorError(f"Invalid start_column value: {self.calender_config['start_column']}. Please check your configuration and try again.")
        if not isinstance(self.calender_config['start_row'], int) or self.calender_config['start_row'] <= 0:
            self.log.critical(f"Invalid start_row value: {self.calender_config['start_row']}. Please check your configuration and try again.")
            raise CalendarXlsxCreatorError(f"Invalid start_row value: {self.calender_config['start_row']}. Please check your configuration and try again.")

        # calculate date cell references and store in the cell_date dictionary:
        row = self.calender_config['start_row']
        col = self.calender_config['start_column']
        for d in range((self.last_date - self.first_date).days + 1):
            date = self.first_date + datetime.timedelta(days=d)
            # calculate cell ref as a touple of (row, col)
            cell_ref = (row, ((col%7) + self.calender_config['start_column'] + + self.first_dow_column))
            # save in dict
            self.cell_date[date] = cell_ref
            #self.log.debug(f"Date {date} mapped to cell {cell_ref}.")






    def get_cell_by_date(self, date: datetime.date) -> cell:
        #get value and return


    def category_row_offset(self, category: str) -> int:
        """Calculate the row offset for a given category based on the category list.
        
        Args:
            category (str): The name of the category for which to calculate the row offset.

        Returns:
            int: The row offset for the given category, which is the index of the category in the category list plus one (to account for date row).
        
        Raises:
            CalendarXlsxCreatorError: If the category is not found in the category list.
        """
        if category not in self.category_list:
            self.log.error(f"Category '{category}' not found in category list.")
            raise CalendarXlsxCreatorError(f"Category '{category}' not found in category list.")
        return self.category_list.index(category) + 1 # add 1 to account for date row

        
        
    def save(self):
        """Save the workbook to the specified path.
        
        Args:
            None
        Retrns:
            None
        
        Raises:
            CalendarXlsxCreatorError: If there is an error saving the workbook.
        
        """
        try:
            self.workbook.save(self.xlsx_path)
            self.log.info(f"Workbook saved successfully at {self.xlsx_path}.")
        except Exception as e:
            self.log.error(f"Failed to save workbook: {e}")
            raise CalendarXlsxCreatorError(f"Failed to save workbook: {e}")



    def mark_date_as_busy(self, date: datetime.date):
        """Mark a specific date as busy with an event in the calendar by changing its style.
        
        Args:
            date (datetime.date): The date to be marked as busy.
            category (str): The category of the event (used for column indexing).
            event_name (str): The name of the event to be displayed in the cell.
        
        Returns:
            Bool: True if the date style was updated, False if it was already marked as busy.
        
        Raises:
            CalendarXlsxCreatorError: If there is an error marking the date as busy, such as invalid date or category.
        
        """
        # This function will be implemented later. It will use the trackers to find the correct cell based on the date and category, and then update the cell with the event name and appropriate styling.
        pass





## algorithm - to be implemetned later:

'''
create worksheet

save


functions:
- load config (on initialisation)
- create worksheet
- save workbook

- update cell style
- add style to workbook

- event_is_multiday

- add event.

---
create matrix.
    create trackers:
        catagory list - index of each catagory releative to the date row
        date index - given a date, return the row and col index
        days of the week 2 column - given a day of the week which column
        column 2 day of the week - given a column, return the day of the week index.
        day of the weeek - iso index
        list of styles created
        
        SELF.
        
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