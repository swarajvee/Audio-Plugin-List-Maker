# Plugin List Generator
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
git clone https://github.com/your-username/plugin-list-generator.git
cd plugin-list-generator

# Step 2: Install dependencies
pip install openpyxl numpy

# Step 3: Running the Script
python plugin_list_generator.py

The script will prompt you to enter the plugin directory paths for VST, VST3, AU, and AAX plugins. Press `Enter` if you don't have a particular type of plugin.
echo "Enter the VST plugin directory path:"
read vst_path
echo "Enter the VST3 plugin directory path:"
read vst3_path
echo "Enter the AU plugin directory path:"
read au_path
echo "Enter the AAX plugin directory path:"
read aax_path

# Execute the script
python plugin_list_generator.py

# Output will be an Excel file created in the current working directory with the filename `plugin list_YYYY-MM-DD_HH-MM-SS.xlsx`
