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




