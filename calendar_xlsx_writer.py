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
    

    
    def __init__(self, config: dict, logger=None, dt=datetime.now().strftime('%Y%m%d_%H%M')):
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