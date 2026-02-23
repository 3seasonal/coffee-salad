class calendar_config:
    """
    This class is used to configure the calendar settings.


    self._config    - dictionary with configuration settings
    self._columns   - column sets used in the calendar
    self._events    - master list events by type used in the calendar
    self._style     - styles used in the calendar



    """
    ''' standard class functions '''
    '''
        config  (dictionary), 
        columns (dictionary), 
        events  (dictionary), 
        styles  (dictionary), 
        utility_class (CalendarUtil object)
    '''

    def __init__(self, config=None, columns=None, events=None, styles=None, utility_class=None, xlsx_class=None):
        """Constructor - Initialize the object"""
        self._config={}
        self._config['config'] = dict(config) if config is not None else {}
        self._config['columns'] = dict(columns) if columns is not None else {}
        self._config['events'] = dict(events) if events is not None else {}
        self._config['styles'] = dict(styles) if styles is not None else {}
        
        # set the utility class to an instatiated object
        self._cu = utility_class if utility_class is not None else None
        print (f'logging enabled in calendar_config: {self.logging_enabled()}')

        self._xlsx = xlsx_class if xlsx_class is not None else None
        print (f'xlsx enabled in calendar_config: {self.xlsx_enabled()}')


    def logging_enabled(self):
        """Check if logging is enabled"""
        return self._cu is not None
    
    def xlsx_enabled(self):
        """Check if xlsx is enabled"""
        return self._xlsx is not None
    
    def __str__(self):
        """String representation for humans"""
        return f"Calendar Config with keys {self._config.keys()}"
    
    def __len__(self):
        """Return length of config"""
        return sum(len(subdict) for subdict in self._config.values() if isinstance(subdict, dict))
    
    def __getitem__(self, key='config'):
        """Allow dictionary-like access: dictionary['key']"""
        if dict == 'config': 
            return self._config.config
        elif dict == 'columns':
            return self._config.columns
        elif dict == 'events':
            return self._config.events
        elif dict == 'styles':
            return self._config.styles
        else:
            return None
    
    " overwrite dictionary configuration "
    def __setitem__(self, key, value):
        """Allow dictionary-like assignment: config['key'] = value"""
        if key in self._config.keys():
            self._config[key] = value
        else:
            raise KeyError(f"Key {key} not found in config dictionary.")

    ''' accessor functions '''
    def get_config(self, key=None):
        """Get configuration value by key"""
        if key is None:
            return self._config.config
        else:
            return self._config.config.get(key, None)

    def get_columns(self, key=None):
        """Get columns value by key"""
        if key is None:
            return self._config.columns
        else:
            return self._config.columns.get(key, None)
    
    def get_events(self, key=None):
        """Get events value by key"""
        if key is None:
            return self._config.events
        else:
            return self._config.events.get(key, None)
        
    def get_styles(self, key=None):
        """Get styles value by key"""
        if key is None:
            return self._config.styles
        else:
            return self._config.styles.get(key, None)


    ''' import / export functions'''
    def import_yaml(self, yaml_file):
        pass

    def export_yaml(self, yaml_file):
        pass

    def import_xlsx(self, xlsx_file):
        pass

    def export_xlsx(self, xlsx_file):
        pass


    " style functions "
    #################
    # Styles - see https://openpyxl.readthedocs.io/en/3.1.3/styles.html
    #styles:
    # style_name:
    #   font:
    #     name: Arial
    #     size: 12
    #     bold: True/False
    #     italic: True/False
    #     strike: True/False
    #     underline: none/single/double 
    #     color: 000000
    #   fill:
    #     fill_type: solid
    #     fgColor: 00FF0000   as ttRRGGBB t=transparency
    #   border:
    #     left:
    #       border_style: {'thin', 'medium', 'dashDot', 'mediumDashed', 'thick', 'dashed', 'slantDashDot', 'mediumDashDot', 'dashDotDot', 'dotted', 'hair', 'double', 'mediumDashDotDot'}
    #       color: 00000000
    #      right:
    #       border_style: {'thin', 'medium', 'dashDot', 'mediumDashed', 'thick', 'dashed', 'slantDashDot', 'mediumDashDot', 'dashDotDot', 'dotted', 'hair', 'double', 'mediumDashDotDot'}
    #       color: 00000000
    #      top:
    #       border_style: {'thin', 'medium', 'dashDot', 'mediumDashed', 'thick', 'dashed', 'slantDashDot', 'mediumDashDot', 'dashDotDot', 'dotted', 'hair', 'double', 'mediumDashDotDot'}
    #       color: 00000000
    #      bottom:
    #       border_style: {'thin', 'medium', 'dashDot', 'mediumDashed', 'thick', 'dashed', 'slantDashDot', 'mediumDashDot', 'dashDotDot', 'dotted', 'hair', 'double', 'mediumDashDotDot'}
    #       color: 00000000
    #   alignment:
    #     horizontal: general/left/center/right/fill/justify/centerContinuous/distributed
    #     vertical: bottom/center/top/justify/distributed
    #     text_rotation: 0
    #     wrap_text: True/False
    #     shrink_to_fit: True/False
    #     indent: 0
    #   number_format: general/0/0.00/0.00%/
    #   protection:
    #     locked: True/False
    #     hidden: True/False

    ''' style path = subdict1/subdict2/value '''
    def get_style(self, style_name, style_path=None):
        """Get styles value by key"""
        if style_path is None:
            return self._config.styles.style_name
        else:
            keys = [list] + style_path.split("/")
            try:
                for key in keys:
                    value = value[key]
                return value
            except (KeyError, TypeError):
                return self._config.styles.get(key, None)
    
    def set_style(self, style_name, style_path=None, value=None):
        """Set styles value by key"""
        if style_path is None:
            self._config.styles.style_name = value
        else:
            keys = [list] + style_path.split("/")
            try:
                for key in keys[:-1]:
                    value = value[key]
                value[keys[-1]] = value
            except (KeyError, TypeError):
                self._config.styles.get(key, None)