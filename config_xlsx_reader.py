import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
import datetime
from typing import Dict, List, Any, Tuple, Optional, Union
import os

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
        self.worksheet = None
        
        # Output dictionaries
        self.calendar = {}
        self.columns = {}
        self.events = {}
        self.styles = {}
        
        # worksheet meta-config
        self.col_value_names = []
        self.col_index = {}
        
        # check if the xlsx file exists
        if not self.xlsx_path:
            raise ValueError("Excel file was not provided.")
        if not os.path.isfile(self.xlsx_path):
            raise FileNotFoundError(f"Excel file '{self.xlsx_path}' does not exist.")
        
        # check if it contains a valid config-main sheet
        self.workbook = openpyxl.load_workbook(self.xlsx_path, data_only=True)
        if "config-main" not in self.workbook.sheetnames:
            self.workbook.close()
            raise ValueError("No main config sheet found.")
        
        else:
            # report the number of sheets loaded
            print(f"loaded workbook '{self.xlsx_path}' with {str(len(self.workbook.sheetnames))} sheets")
         
    
    def read_config(self) -> Tuple[Dict, Dict, Dict, Dict]:
        """
        Reads the configuration from the Excel file and returns the four dictionaries.
        Follows the configuration tree, starting from config-meta, then config-main, and only parses referenced config sheets (ignoring config-meta as a config).
        Returns:
            Tuple containing (calendar, columns, events, styles) dictionaries
        """       
        # parse the config-main sheet and follow the configuration tree
        parsed_config = self._parse_worksheet_(self, "config-main")
        
        return parsed_config.calendar, parsed_config.columns, parsed_config.events, parsed_config.styles



    def _parse_worksheet_(self, worksheet_name) -> Dict:
        """_summary_
        This function parses a single worksheet from the configuriation and returns 
        a dictionary representation of its contents.
        Function is recursive and will follow config references to other sheets.
        will check the first two rows for meta-config and header row.
        This will be used to configure how the rest of the sheet is parsed.
        Args:
            worksheet_name (string): the name of the worksheet to parse
        Returns:
            Dict: a dictionary representation of the worksheet contents
        THrows:
            ValueError: if the worksheet does not exist
        """

        
        # get the worksheet and meta-config
        worksheet, col_index, col_value_names = self._load_worksheet_metaconfig_(worksheet_name)
        
        # iniatialize output dictionary
        output_dict = {}
        
        # Build a dictionary mapping column names to their indices
        col_names = {
            'param': col_index.get('param'),
            'type': col_index.get('type')
        }
        for value_name in col_value_names:
            col_names[value_name] = col_index.get(value_name)
                
        # Parse the content in the sheet based on the meta-config
        for row in worksheet.iter_rows(min_row=3):
            
            #if the param cell is empty or commented out, skip the row
            if (row[col_names['param'] - 1].value is not None) and (row[col_names['param'] - 1].value.strip() != "") and not (row[col_names['param'] - 1].value.strip().startswith("#")):
                
                param = row[col_names['param'] - 1].value.strip().lower()
                param_type = row[col_names['type'] - 1].value.strip().lower() if row[col_names['type'] - 1].value else "string"
                
                # Handle different prinitive types
                if param_type == "string":
                    output_dict[param] = row[col_names['value'] - 1].value.strip()
                elif param_type == "bool":
                    output_dict[param] = True if row[col_names['value'] - 1].value.strip().lower() in ["true", "yes", "1"] else False
                elif param_type == "int":
                    try:
                        output_dict[param] = int(row[col_names['value'] - 1].value.strip())
                    except (ValueError, TypeError):
                        raise ValueError(f"Invalid integer value for parameter '{param}' in sheet '{worksheet_name}' Row {row[0].row}")
                elif param_type == "float":
                    try:
                        output_dict[param] = float(row[col_names['value'] - 1].value.strip())
                    except (ValueError, TypeError):
                        raise ValueError(f"Invalid float value for parameter '{param}' in sheet '{worksheet_name}' Row {row[0].row}")
                    output_dict[param] = float(row[col_names['value'] - 1].value.strip())
                elif param_type == "date":
                    output_dict[param] = self._parse_date(row[col_names['value'] - 1].value.strip())
                    
                # Handle different complex types  
                elif param_type == "list":
                    sub_dict = {}
                    
                    
                    # unpack the actual name and type form the col_value_names
                    # col_value_names.remove('value') #keep the value column as it is will act as an alias
                    for vc in col_value_names:
                        if "_" in vc:
                            vc_type = vc.split("_")[-1]
                            vc = vc[:-len(vc_type)]
                        else:
                            vc_type = "string"
                        
                        #...
                        # handle conversion of each value to its type
                        #... and set to dict
                        sub_dict[vc] = row[col_names[vc] - 1].value.strip() if row[col_names[vc] - 1].value else ""
                        
                    # add sub-dict to output dict
                    output_dict[param] = sub_dict
                        
                        
                
                ## Handle config reference recursion
                elif param_type == "config":
                    # Recursively parse the referenced config sheet
                    ref_sheet_name = row[col_names['value'] - 1].value.strip()
                    if ref_sheet_name and ref_sheet_name.strip() != "":
                        output_dict[param] = self._parse_worksheet_(ref_sheet_name.strip())
                    else:
                        raise ValueError(f"Empty config reference for parameter '{param}' in sheet '{worksheet_name}' Row {row[0].row}")
                else:
                    raise ValueError(f"Unknown type '{param_type}' for parameter '{param}' in sheet '{worksheet_name}'")
            
        # finished - turn in home work
        return output_dict
            
            
'''
NOTE TO SELF:
need to update the config to include _{typw} type for all value columns that are in lists
eg start_date 


'''
    


    def _get_worksheet_(self, sheet_name: str) -> openpyxl.worksheet:
        
        
        """Check if a worksheet exists in the workbook."""
        if sheet_name is None or sheet_name.strip() == "":
            raise ValueError("Sheet name must be a non-empty string.")
        if sheet_name not in self.workbook.sheetnames:
            raise ValueError(f"Worksheet '{sheet_name}' does not exist in the workbook containing: ({', '.join(self.workbook.sheetnames)})")
        
        return self.workbook[sheet_name]   



      
    def _get_worksheet_metaconfig_(self, sheet_name) -> Tuple[openpyxl.worksheet, dict, list]:
        
        # reset col_value_names
        col_value_names = []
        col_index = {}
        worksheet = self._get_worksheet_(sheet_name)
                
        # read first row to get column names
        first_row = [cell.value for cell in worksheet[1]]
        
        if ("param" not in first_row) or ("type" not in first_row) or ("value" not in first_row):
            raise ValueError(f"Missing required columns in header: {first_row}")
        
        # convert list to dict of column name to 1 based index
        col_index = {col_name: idx + 1 for idx, col_name in enumerate(first_row)} 
        # get value columns
        col_value_names = self.worksheet.cell(row=2, column=col_index["value"]).value.split(",")
        
        return worksheet, col_index, col_value_names
        
        
        
        
        
        
        
    
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
    
    ########################################
    
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
