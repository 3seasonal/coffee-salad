import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
import datetime
from typing import Dict, List, Any, Tuple, Optional, Union

class ConfigXlsxReader:
    """
    Reads calendar configuration from an Excel file.
    Extracts configuration into four dictionaries:
    - calendar: Calendar configuration (start/end dates, etc.)
    - columns: Calendar column structures
    - events: Event types and entries
    - styles: Styles used in the calendar
    """
    
    def __init__(self, xlsx_path: str):
        """Initialize the reader with the path to the Excel file."""
        self.xlsx_path = xlsx_path
        self.workbook = None
        
        # Output dictionaries
        self.calendar = {}
        self.columns = {}
        self.events = {}
        self.styles = {}
        
        # For tracking meta configurations
        self.meta_config = {}
    
    def read_config(self) -> Tuple[Dict, Dict, Dict, Dict]:
        """
        Reads the configuration from the Excel file and returns the four dictionaries.
        
        Returns:
            Tuple containing (calendar, columns, events, styles) dictionaries
        """
        # Open the workbook
        self.workbook = openpyxl.load_workbook(self.xlsx_path, data_only=True)
        
        # Find all config worksheets and process them
        config_sheets = [sheet for sheet in self.workbook.sheetnames if sheet.startswith("config-")]
        
        # Process meta sheet first if it exists
        if "config-meta" in config_sheets:
            self._process_meta_sheet("config-meta")
            config_sheets.remove("config-meta")
        
        # Find main config sheet (first non-meta sheet)
        main_config = next(iter(config_sheets), None)
        if main_config:
            self._process_main_config(main_config)
            config_sheets.remove(main_config)
        
        # Process styles if exists
        if "config-style" in config_sheets:
            self._process_style_sheet("config-style")
            config_sheets.remove("config-style")
        
        # Process remaining sheets based on references from the main config
        for sheet_name in config_sheets:
            self._process_config_sheet(sheet_name)
        
        # Close workbook
        self.workbook.close()
        
        return self.calendar, self.columns, self.events, self.styles
    
    def _process_meta_sheet(self, sheet_name: str) -> None:
        """Process the meta configuration sheet."""
        sheet = self.workbook[sheet_name]
        header_row = self._find_header_row(sheet)
        
        if not header_row:
            return
        
        headers = self._get_row_values(sheet, header_row)
        
        # Find columns for type, param, and value
        type_idx = headers.index('type') if 'type' in headers else None
        param_idx = headers.index('param') if 'param' in headers else None
        value_idx = headers.index('value') if 'value' in headers else None
        
        if not all([type_idx is not None, param_idx is not None, value_idx is not None]):
            return
            
        # Process the meta-config row (first data row after headers)
        meta_row = self._get_row_values(sheet, header_row + 1)
        if len(meta_row) <= max(type_idx, param_idx, value_idx):
            return
            
        # Store meta-config for parsing other sheets
        self.meta_config = {
            "type_col": type_idx,
            "param_col": param_idx,
            "value_col": value_idx,
            # Other meta config as needed
        }
    
    def _process_main_config(self, sheet_name: str) -> None:
        """Process the main configuration sheet and build the initial configuration."""
        sheet = self.workbook[sheet_name]
        config_data = self._parse_config_sheet(sheet)
        
        # Organize data into the appropriate dictionaries based on the config structure
        for item in config_data:
            if item.get("type") == "calendar":
                self.calendar = self._process_calendar_config(item)
            elif item.get("type") == "columns":
                self.columns["config"] = item
                if "sheet" in item:
                    self._process_columns_sheet(item["sheet"])
            elif item.get("type") == "events":
                self.events["config"] = item
                if "sheet" in item:
                    self._process_events_sheet(item["sheet"])
            elif item.get("type") == "styles":
                self.styles["config"] = item
                # Styles are usually processed separately through _process_style_sheet
    
    def _process_calendar_config(self, config: Dict) -> Dict:
        """Process calendar-specific configuration."""
        # Extract and transform calendar configuration
        calendar_config = {}
        
        # Process based on config structure from README
        if "start_date" in config:
            calendar_config["start_date"] = config["start_date"]
        if "end_date" in config:
            calendar_config["end_date"] = config["end_date"]
        # Add other calendar-specific properties
        
        return calendar_config
    
    def _process_columns_sheet(self, sheet_name: str) -> None:
        """Process the columns configuration sheet."""
        if sheet_name not in self.workbook.sheetnames:
            return
            
        sheet = self.workbook[sheet_name]
        column_data = self._parse_config_sheet(sheet)
        
        # Process column class list with metadata
        column_classes = []
        for item in column_data:
            if item.get("type") == "column-class":
                column_classes.append(item)
            # Process other column-related configurations
        
        self.columns["classes"] = column_classes
        
        # Process any sub-configurations for columns if needed
        for col_class in column_classes:
            if "sheet" in col_class:
                sub_sheet_name = col_class["sheet"]
                if sub_sheet_name in self.workbook.sheetnames:
                    # Process sub-configuration
                    sub_data = self._parse_config_sheet(self.workbook[sub_sheet_name])
                    col_class["sub_config"] = sub_data
    
    def _process_events_sheet(self, sheet_name: str) -> None:
        """Process the events configuration sheet."""
        if sheet_name not in self.workbook.sheetnames:
            return
            
        sheet = self.workbook[sheet_name]
        event_data = self._parse_config_sheet(sheet)
        
        # Process event class list with metadata
        event_classes = []
        event_entries = []
        
        for item in event_data:
            if item.get("type") == "event-class":
                event_classes.append(item)
            elif item.get("type") == "event":
                event_entries.append(item)
            # Process other event-related configurations
        
        self.events["classes"] = event_classes
        self.events["entries"] = event_entries
    
    def _process_style_sheet(self, sheet_name: str) -> None:
        """Process the styles configuration sheet."""
        if sheet_name not in self.workbook.sheetnames:
            return
            
        sheet = self.workbook[sheet_name]
        style_data = self._parse_config_sheet(sheet)
        
        cell_styles = {}
        border_styles = {}
        
        for item in style_data:
            if item.get("type") == "style-cell":
                # Extract cell style properties
                style_name = item.get("param")
                if style_name:
                    cell = self._get_cell_by_value(sheet, item.get("value"))
                    if cell:
                        cell_styles[style_name] = self._extract_cell_style(cell)
            elif item.get("type") == "style-border":
                # Extract border style properties
                style_name = item.get("param")
                if style_name:
                    cell = self._get_cell_by_value(sheet, item.get("value"))
                    if cell:
                        border_styles[style_name] = self._extract_border_style(cell)
        
        self.styles["cell_styles"] = cell_styles
        self.styles["border_styles"] = border_styles
    
    def _process_config_sheet(self, sheet_name: str) -> None:
        """Process any other configuration sheet based on its content."""
        # This would be implemented based on the specific needs of other config sheets
        pass
    
    def _parse_config_sheet(self, sheet) -> List[Dict]:
        """
        Parse a configuration sheet and return a list of configuration items.
        
        Args:
            sheet: The worksheet to parse
            
        Returns:
            List of dictionaries with configuration items
        """
        result = []
        header_row = self._find_header_row(sheet)
        
        if not header_row:
            return result
            
        headers = self._get_row_values(sheet, header_row)
        
        # Find important column indices
        type_idx = headers.index('type') if 'type' in headers else self.meta_config.get("type_col")
        param_idx = headers.index('param') if 'param' in headers else self.meta_config.get("param_col")
        value_idx = headers.index('value') if 'value' in headers else self.meta_config.get("value_col")
        
        if not all([type_idx is not None, param_idx is not None, value_idx is not None]):
            return result
        
        # Process meta-config row if exists
        meta_row_idx = header_row + 1
        meta_row = self._get_row_values(sheet, meta_row_idx)
        
        if len(meta_row) > max(type_idx, param_idx, value_idx):
            meta_config = {
                "value_columns": self._parse_value_columns(meta_row[value_idx]) if value_idx < len(meta_row) else None
            }
        else:
            meta_config = {}
            
        # Process data rows
        for row_idx in range(header_row + 2, sheet.max_row + 1):
            row_values = self._get_row_values(sheet, row_idx)
            
            # Skip empty rows or comment rows
            if not row_values or (row_values and row_values[0].startswith('#')):
                continue
                
            if len(row_values) <= max(type_idx, param_idx, value_idx):
                continue
                
            item_type = row_values[type_idx] if type_idx < len(row_values) else None
            param_name = row_values[param_idx] if param_idx < len(row_values) else None
            
            # Skip if no type or param
            if not item_type or not param_name:
                continue
                
            # Create config item
            config_item = {
                "type": item_type,
                "param": param_name
            }
            
            # Process value based on type
            if item_type == "int":
                config_item[param_name] = int(row_values[value_idx]) if value_idx < len(row_values) else None
            elif item_type == "string":
                config_item[param_name] = row_values[value_idx] if value_idx < len(row_values) else ""
            elif item_type == "date":
                config_item[param_name] = self._parse_date(row_values[value_idx]) if value_idx < len(row_values) else None
            elif item_type == "list":
                # Process list values based on meta_config
                if meta_config.get("value_columns"):
                    list_values = []
                    for col_idx in meta_config["value_columns"]:
                        if col_idx < len(row_values):
                            list_values.append(row_values[col_idx])
                    config_item[param_name] = list_values
            elif item_type == "config":
                # Link to another sheet
                config_item["sheet"] = row_values[value_idx] if value_idx < len(row_values) else None
            elif item_type in ["style-cell", "style-border"]:
                config_item["value"] = row_values[value_idx] if value_idx < len(row_values) else None
                
            result.append(config_item)
            
        return result
    
    def _find_header_row(self, sheet) -> Optional[int]:
        """Find the header row in the sheet."""
        for row_idx in range(1, min(10, sheet.max_row + 1)):  # Check first 10 rows
            row_values = self._get_row_values(sheet, row_idx)
            if 'type' in row_values and 'param' in row_values:
                return row_idx
        return None
    
    def _get_row_values(self, sheet, row_idx: int) -> List[str]:
        """Get values from a row, converting to strings."""
        values = []
        for cell in sheet[row_idx]:
            if cell.value is not None:
                values.append(str(cell.value).strip())
            else:
                values.append("")
        return values
    
    def _parse_value_columns(self, value_str: str) -> List[int]:
        """Parse value column string into column indices."""
        if not value_str or value_str.lower() == 'value':
            # Default to just the value column
            return []
            
        # Parse comma-separated column references
        columns = []
        for col_ref in value_str.split(','):
            col_ref = col_ref.strip()
            if col_ref:
                try:
                    # Handle both numeric indices and column letters (A, B, etc.)
                    if col_ref.isalpha():
                        columns.append(column_index_from_string(col_ref) - 1)  # 0-based index
                    else:
                        columns.append(int(col_ref) - 1)  # 0-based index
                except (ValueError, KeyError):
                    pass
        return columns
    
    def _parse_date(self, date_str: str) -> Optional[datetime.date]:
        """Parse a date string into a date object."""
        if not date_str:
            return None
            
        # First check if it's already a datetime
        if isinstance(date_str, datetime.datetime):
            return date_str.date()
            
        # Try various date formats
        date_formats = [
            "%Y-%m-%d",
            "%d/%m/%Y",
            "%m/%d/%Y",
            "%d-%m-%Y",
            "%d %b %Y",
            "%d %B %Y"
        ]
        
        for fmt in date_formats:
            try:
                return datetime.datetime.strptime(date_str, fmt).date()
            except ValueError:
                continue
                
        return None
    
    def _get_cell_by_value(self, sheet, value: str) -> Optional[object]:
        """Find a cell with the given value in the sheet."""
        if not value:
            return None
            
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value == value:
                    return cell
        return None
    
    def _extract_cell_style(self, cell) -> Dict:
        """Extract style properties from a cell."""
        style = {}
        
        # Font properties
        if cell.font:
            style["font"] = {
                "name": cell.font.name,
                "size": cell.font.size,
                "bold": cell.font.bold,
                "italic": cell.font.italic,
                "color": cell.font.color.rgb if cell.font.color else None
            }
        
        # Fill properties
        if cell.fill:
            style["fill"] = {
                "type": cell.fill.fill_type,
                "start_color": cell.fill.start_color.rgb if cell.fill.start_color else None,
                "end_color": cell.fill.end_color.rgb if cell.fill.end_color else None
            }
        
        # Alignment properties
        if cell.alignment:
            style["alignment"] = {
                "horizontal": cell.alignment.horizontal,
                "vertical": cell.alignment.vertical,
                "wrap_text": cell.alignment.wrap_text,
                "text_rotation": cell.alignment.text_rotation
            }
        
        return style
    
    def _extract_border_style(self, cell) -> Dict:
        """Extract border style properties from a cell."""
        border_style = {}
        
        if not cell.border:
            return border_style
            
        # Process each side of the border
        for side in ["top", "right", "bottom", "left"]:
            border_side = getattr(cell.border, side)
            if border_side:
                border_style[side] = {
                    "style": border_side.style,
                    "color": border_side.color.rgb if border_side.color else None
                }
                
        return border_style
