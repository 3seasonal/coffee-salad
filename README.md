# coffee-salad

**Pythonic Excel calendar generator**

A Python application to create a calendar in Microsoft Excel.

- Script configuration is stored in a YAML file.
- Calendar configuration is stored in an `.xlsx` file and details:
  - The calendar configuration, including start and end dates.
  - The calendar column structures.
  - The types of events entered into the calendar.
  - The actual events entered, including start/end dates and type.
  - The styles used in the calendar.

## Calendar_template

columns and rows
- First row contains the column headders
- any empty rows are ignored, for readability only 
- any rows starting with # in the first populated column can be considered empty and ignored
- columns that don't have a header are ignored

types
- each parameter is asigned a type, these are used to contextualise values and define the configuration structure
- types are defined in the meta-config 
- meta-config type params provide the metadata for the proceeding configuration in the sheet
- config type params provide a link to the next sub-configuration worksheet
- meta types point to non-parsable metadata sheets and should not be read
- int,string,date describe data types and how they should be parsed
- list type indicates that the parapmeter is a list and will be read as a 2d dictionary

meta-config
- the metaconfig is usually the first row of values after the headers
- this row describes how to parse the proceeding information

param column
- param column defines the parameters name
- when the value type is a list, the meta-config value will be used as the matrix/list/dictionary name

type column
- each parameter is asigned a type, these are used to contextualise values and define the configuration structure
- types are defined in the meta-config 
- meta-config type params provide the metadata for the proceeding configuration in the sheet
- config type params provide a link to the next sub-configuration worksheet
- meta types point to non-parsable metadata sheets and should not be read
- int,string,date describe data types and how they should be parsed
- list type indicates that the parapmeter is a list and the restcolumn header
- style-cell uses the styles the value column (colour, bg, angle etc) for parsign style properties - borders are ignored
- style-border uses the styles the value column borders for parsign the border style and sides

value column
- in the meta-config the value column defines which columns to use as values
- single value parameters usually say 'value' in the sheets meta-config row, else a list of columns to consider the as a list (row) value
- trailing commers can be ignored

other columns
- unlsess specified in the value column, all other columns are meta and can be ignored
- meta columns (usually but not always) include description, notes, ref

Worksheets are used to store individual configurations
- worsheet names must start with "config-"
- the "config-meta" sheet defines lists that may be used later. it does not 
- first non-meta worksheet is the starting index or "main config" sheet, usualy called config-main
- from the main config, the condiguration is mapped out in a tree structure


## Calendar_template structure:
```mermaid
flowchart LR
    A[configuration index<br/>and metadata]:::orange --> B[calendar]:::grey
    A --> C[columns]:::grey
    A --> D[events]:::grey
    A --> E[styles]:::grey

    B --> F[table of values]:::blue

    C --> G[column class list<br/>with metadata]:::orange
    G --> H[column class table]:::green
    H -.-> I[column class subtable]:::green

    D --> J[event class list<br/>with metadata]:::orange
    J --> K[event list table]:::green

    E --> L[border styles<br/>metadata and list]:::orange
    L --> M[styled cells]:::purple
    L --> N[bordered cells]:::purple

    classDef orange fill:#f9d29d,stroke:#d4a373,stroke-width:1px;
    classDef grey fill:#e5e5e5,stroke:#999,stroke-width:1px;
    classDef green fill:#b6e3b6,stroke:#4c9141,stroke-width:1px;
    classDef purple fill:#d5b8e5,stroke:#7a4c91,stroke-width:1px;
    classDef blue fill:#c9daf8,stroke:#3c78d8,stroke-width:1px;




## Calendar

The output calendar is saved as a sheet in an xlsx file.


## Classes

Includes the following classes:

- `main`
- `calendar_creator`
- `config_xlsx_reader`  
  Reads the structure, content, and styles of a calendar from an `.xlsx` (config) file.
- `config_xlsx_writer`  
  Saves the structure, content, and styles of a calendar to an `.xlsx` (config) file.
- `calendar_util`  
  Holds logging, script config loading, and other utility functions.


