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
import sys
import os
from datetime import datetime
import yaml

#### set config file ####
config_source_name = "calendar_config.yaml"
#########################
script_dir = os.path.dirname(os.path.abspath(__file__))
script_config_file = os.path.join(script_dir, config_source_name)


# Ensure the project root (script directory) is on sys.path so local modules can be imported
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# import local modules
from calendar_util import CalendarUtility #for logging and other utilities
#from calendar_xlsx_writer import calendarXlsxCreator #for writing the calendar xlsx file



# currently just a testing script
def main(script_config_file=script_config_file):
    print("Welcome to the Coffee-Salad Calendar Application!")
    
    # get the path to the folder this script is in:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"Script is located in: {script_dir}")
    print(f"Using config file: {script_config_file}")

    # read the yaml config_file:
    with open(script_config_file, 'r') as f:
        script_config = yaml.safe_load(f)
    # get yaml config value: for config_source_name
    calendar_config_file = script_config["calendar_config_file"]
    # if the file name is not an absolute path, assume it is relative to the script directory:
    if not os.path.isabs(calendar_config_file):
        calendar_config_file = os.path.join(script_dir, calendar_config_file)
    print(f"Calendar source name: {calendar_config_file}")

    # set up logging
    ts = datetime.now().strftime('%Y%m%d_%H%M%S')
    ds = datetime.now().strftime('%Y%m%d')
    logpath = os.path.join(script_dir, 'log')
    logname = f'calendar_app_{ts}.log'

    # load utils
    util = CalendarUtility(logpath=logpath, logname=logname)
    log = util.get_logger()

      
    if not os.path.exists(calendar_config_file):
        raise FileNotFoundError(f"Template file not found: {calendar_config_file}")
    
    from config_xlsx_reader import ConfigXlsxReader
    xlsx_reader = ConfigXlsxReader(xlsx_path=calendar_config_file, logger=log)
    
    calendar, columns, events, styles = xlsx_reader.read_config()
    
    # pretty print dictionaries:
    import pprint
    pp = pprint.PrettyPrinter(indent=4)
    log.info("Calendar Configuration:")
    pp.pprint(calendar)
    log.info("Columns Configuration:")
    pp.pprint(columns)
    log.info("Events Configuration:")
    pp.pprint(events)
    log.info("Styles Configuration:")
    pp.pprint(styles)
    
    # initialise the writer
    #xlsx_writer = calendarXlsxCreator(
        # config= {
        #     'calendar': calendar,
        #     'columns': columns,
        #     'events': events,
        #     'styles': styles    
        #     },
        # logger=log,
        # ts=ts)

    
    
    
if __name__ == "__main__":
    main(script_config_file) 