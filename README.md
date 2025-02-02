# Audio-Plugin-list-Maker
This Python script generates an Excel file listing all plugins from specified directories for VST, VST3, AU, and AAX formats. 
It categorizes them by plugin type and also creates a separate sheet for common plugins across formats along with their versions.

# Features:
- Scans specified directories for plugins in VST, VST3, AU, and AAX formats.
- Generates an Excel file with detailed plugin lists.
- Creates a second sheet with common plugins across different formats and their versions.

# Prerequisites:
1. Python 3.x
2. `openpyxl` and `numpy` libraries

# Step 1: Clone the repository
```sh
git clone git@github.com:swarajvee/Audio-Plugin-List-Maker.git
cd Audio-Plugin-List-Maker
```

# Step 2: Install dependencies
`pip install -r requirements.txt`

# Step 3: Running the Script
`python plugin_list_generator.py`

The script will prompt you to enter the plugin directory paths for VST, VST3, AU, and AAX plugins. Press `Enter` if you don't have a particular type of plugin.
Output will be an Excel file created in the current working directory with the filename `plugin list_YYYY-MM-DD_HH-MM-SS.xlsx`
