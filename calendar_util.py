import os
from datetime import datetime
import logging
import sys

# This file is part of the coffee-salad calandar package.
'''
handles logging

'''

# Calendar Utility Class
class CalendarUtility:
    def __init__(self, logpath=None, logname=None):
        # Constructor - Initialize the object
        self._datetime = datetime.now().strftime('%Y%m%d_%H%M')
        self._logpath = str(logpath) if logpath is not None else os.join(os.path.abspath(__file__),'log')
        self._logname = str(logname) if logname is not None else {}
        self._logfile = os.path.join(self._logpath, self._logname)
        
        # check if logpath exists
        if not os.path.exists(self._logpath):
            os.makedirs(self._logpath)

        # set up logging
        self.log_setup(logging.INFO, logging.DEBUG)

        self.log(f'initialised utility class', level='INFO')
        self.log(f'log file created at {self._logfile}', level='INFO')  


    def log_setup(self, console_level=logging.INFO, file_level=logging.DEBUG):
    
        # Create logger
        self.logger = logging.getLogger(self._logname)
        self.logger.setLevel(logging.DEBUG)  # Set to lowest level to capture everything
        print("debug")
        
        # Console handler for INFO and above
        console_handler = logging.StreamHandler(sys.stdout)
        print(console_level)
        console_handler.setLevel(console_level)
        console_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        console_handler.setFormatter(console_formatter)

        # File handler for DEBUG and above
        print (f'initialising logfile {self._logfile}')
        file_handler = logging.FileHandler(self._logfile)
        file_handler.setLevel(logging.DEBUG)  # Corrected from logging.debug to logging.DEBUG
        file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(file_formatter)

        # Add handlers to logger
        self.logger.addHandler(console_handler)
        self.logger.addHandler(file_handler)

        # Example usage
        self.logger.info(f'writing log to {self._logfile}')
        self.logger.debug(f'set file-logging to {file_level} level')
        self.logger.info(f'set console-logging to {console_level} level')

        # set default openXLS colors
        self.COLOR_INDEX = [ # Default Color Index as per 18.8.27 of ECMA Part 4
            '00000000', '00FFFFFF', '00FF0000', '0000FF00', '000000FF', #0-4
            '00FFFF00', '00FF00FF', '0000FFFF', '00000000', '00FFFFFF', #5-9
            '00FF0000', '0000FF00', '000000FF', '00FFFF00', '00FF00FF', #10-14
            '0000FFFF', '00800000', '00008000', '00000080', '00808000', #15-19
            '00800080', '00008080', '00C0C0C0', '00808080', '009999FF', #20-24
            '00993366', '00FFFFCC', '00CCFFFF', '00660066', '00FF8080', #25-29
            '000066CC', '00CCCCFF', '00000080', '00FF00FF', '00FFFF00', #30-34
            '0000FFFF', '00800080', '00800000', '00008080', '000000FF', #35-39
            '0000CCFF', '00CCFFFF', '00CCFFCC', '00FFFF99', '0099CCFF', #40-44
            '00FF99CC', '00CC99FF', '00FFCC99', '003366FF', '0033CCCC', #45-49
            '0099CC00', '00FFCC00', '00FF9900', '00FF6600', '00666699', #50-54
            '00969696', '00003366', '00339966', '00003300', '00333300', #55-59
            '00993300', '00993366', '00333399', '00333333', '00000000', #60-64 - Assueme foreground colour (64) is black
        ]
        
        
    def get_color_at_index(self, index=None):
        '''
        Return the color at the specified index from the COLOR_INDEX list. 
        If no index is provided, return the entire COLOR_INDEX list.
        See: https://openpyxl.readthedocs.io/en/3.1/styles.html#colors for more details on openpyxl color handling.
        Note: The alpha value refers in theory to the transparency of the colour but this is not relevant for cell styles. The default of 00 will prepended to any simple RGB value
        Args:
            index (int, optional): The index of the color to retrieve from the COLOR_INDEX list. If None, the entire COLOR_INDEX list will be returned. Defaults to None.
        Returns:
            str or [str]: A hex color string in the format "AARRGGBB" - note the alpha chanel refers to the transparency of the colour but this is not relevant for cell styles. The default of 00 will prepended to any simple RGB value.
        Raises:
            none.
        
        '''    
        # return the color at the index
        if index is not None and 0 <= index < len(self.COLOR_INDEX):
            return self.COLOR_INDEX[index]
        
        # otherwise return the whole list
        if not index:
            return self.COLOR_INDEX
        
    

    def get_logger(self):
        return self.logger

    ## Log a message to output
    # - Note this is only accessable if accessing logging through the wrapper.
    # - If accessing the logger directly (via the ge_logger), use the standard logging methods (e.g. logger.info(), logger.debug(), etc.)
    def log(self, message, level='INFO'):
        # Log a message to the log file
        if 'critical' in level.lower():
            self.logger.critical(message)
        elif 'error' in level.lower():
            self.logger.error(message)
        elif 'warning' in level.lower():
            self.logger.warning(message)
        elif 'info' in level.lower():
            self.logger.info(message)
        elif 'debug'in level.lower():
            self.logger.debug(message)
        else:
            self.logger.info(message)




