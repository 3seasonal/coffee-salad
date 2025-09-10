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

# currently just a testing script
def main():
    print("Welcome to the Coffee-Salad Calendar Application!")
    
    # get the path to the folder this script is in:
    import os
    script_dir = os.path.dirname(os.path.abspath(__file__))
    print(f"Script is located in: {script_dir}")
    
    testxlsx = os.path.join(script_dir, 'calendar_template.xlsx')
    
    if not os.path.exists(testxlsx):
        raise FileNotFoundError(f"Template file not found: {testxlsx}")
    
    from config_xlsx_reader import ConfigXlsxReader
    xlsx_reader = ConfigXlsxReader(testxlsx)
    
    calendar, columns, events, styles = xlsx_reader.read_config()
    
    # pretty print dictionaries:
    import pprint
    pp = pprint.PrettyPrinter(indent=4)
    print("Calendar Configuration:")
    pp.pprint(calendar)
    print("Columns Configuration:")
    pp.pprint(columns)
    print("Events Configuration:")
    pp.pprint(events)
    print("Styles Configuration:")
    pp.pprint(styles)
    
    
    
    
if __name__ == "__main__":
    main() 