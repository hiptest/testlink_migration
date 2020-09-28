# testlink_migration

This script performs a migration from a Testlink project to CS project, in order to build the same scheme as the spreadsheet template (.xlxs) available in the app.

### 1. Install dependencies
- Python 3
- lxml
- bs4
- openpyxl

### 2. Export your TestLink project as .XML format

### 3. Execute the python script

`python <migration_script.py> –i <path_to_exported_testlink_project.xml>`

### 4. The script will generate a .xlsx file

### 5. In CS, import this generated .xlsx file 
