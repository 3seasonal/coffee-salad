#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""Calendar xlsx main
    Coffee-Salad calendar package

    This script is the main entry point for the Coffee-Salad calendar application.
    It initializes the application, handles command-line arguments, and coordinates
        
    Author:         dillonj
    Created:        2025-08-29
    Description:    Class for kicking it all off

"""
from calendar_util import CalendarUtil
from datetime import datetime
import os

# currently just a testing script
def main():
    print("Welcome to the Coffee-Salad Calendar Application!")
    
    # get the path to the folder this script is in:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"Script is located in: {script_dir}")
    
    dt = datetime.now().strftime('%Y%m%d_%H%M')
    logpath = os.path.join(script_dir, 'log')
    logname = f'calendar_app_{dt}.log'

    # load utils
    util = CalendarUtil(logpath=logpath, logname=logname)
    log = util.get_logger()

    testxlsx = os.path.join(script_dir, 'calendar_template.xlsx')
    
    if not os.path.exists(testxlsx):
        raise FileNotFoundError(f"Template file not found: {testxlsx}")
    
    from config_xlsx_reader import ConfigXlsxReader
    xlsx_reader = ConfigXlsxReader(xlsx_path=testxlsx, logger=log)
    
    calendar, columns, events, styles = xlsx_reader.read_config()
    
    # pretty print dictionaries:
    import pprint
    pp = pprint.PrettyPrinter(indent=4)
    log.log("Calendar Configuration:")
    pp.pprint(calendar)
    log.log("Columns Configuration:")
    pp.pprint(columns)
    log.log("Events Configuration:")
    pp.pprint(events)
    log.log("Styles Configuration:")
    pp.pprint(styles)
    
    
    
    
if __name__ == "__main__":
    main() 