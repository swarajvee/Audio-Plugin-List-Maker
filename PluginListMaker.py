import os
import openpyxl
from datetime import datetime
import numpy as np

def ExcelMaker(vst_list, au_list, vst3_list, aax_list):
    """Makes excel file"""

    wb = openpyxl.Workbook()
    sheet = wb.active
    sheet.title = 'My Plugin List'
    headers = ['Sl.No','VST3 Plugins', 'VST Plugins', 'AU Plugins', 'AAX Plugins']
    sheet.append(headers)

    sl_no = np.arange(1,len(max(vst3_list, vst_list, au_list, aax_list, key=len))+1, 1)
    for o, no in enumerate(sl_no, start=2):
        sheet.cell(row=o, column=1, value=no)
                      
    if vst3_list:
        for i, vst3 in enumerate(vst3_list, start=2):
            sheet.cell(row=i, column=2, value=vst3.replace('.vst3', ''))
    if vst_list:
        #c2 = sheet.cell(row=1, column=2, value='VST Plugins')
        for j, vst in enumerate(vst_list, start=2):
            sheet.cell(row=j, column=3, value=vst.replace('.vst', ''))
    if au_list:
        #c3 = sheet.cell(row=1, column=3, value='AU Plugins')
        for k, au in enumerate(au_list, start=2):
            sheet.cell(row=k, column=4, value=au.replace('.component', ''))
    if aax_list:
        #c4 = sheet.cell(row=1, column=4, value='AAX Plugins')
        for l, aax in enumerate(aax_list, start=2):
            sheet.cell(row=l, column=5, value=aax.replace('.aaxplugin', ''))
    
    sheet.row_dimensions[1].height = 48
    sheet.column_dimensions['B'].width = 35
    sheet.column_dimensions['C'].width = 35
    sheet.column_dimensions['D'].width = 35
    sheet.column_dimensions['E'].width = 40

    #sheet 2 things
    common_plugins_sheet = wb.create_sheet(title='Common Plugins')
    headers_common = ['Sl. No','Common Plugins', 'Versions']
    common_plugins_sheet.append(headers_common)

    combined_plugins = set(vst3_list) | set(vst_list) | set(au_list) | set(aax_list)
    common_plugins = []
    plugin_versions = []

    for plugin in combined_plugins:
        versions = []
        if vst3_list:
            if plugin in vst3_list:
                versions.append('VST3')
        if vst_list:        
            if plugin in vst_list:
                versions.append('VST')
        if au_list:
            if plugin in au_list:
                versions.append('AU')
        if aax_list:
            if plugin in aax_list:
                versions.append('AAX')
        if len(versions)>1:
            common_plugins.append(plugin)
            plugin_versions.append(','.join(versions))

    sl_no_common = np.arange(1,len(common_plugins)+1, 1)
    for p, no in enumerate(sl_no_common, start=2):
        common_plugins_sheet.cell(row=p, column=1, value=no)

    for m, (common_plugin, version) in enumerate(zip(common_plugins, plugin_versions), start=2):
        common_plugins_sheet.cell(row=m, column=2, value=common_plugin)
        common_plugins_sheet.cell(row=m, column=3, value=version)

    
    common_plugins_sheet.row_dimensions[1].height = 48
    common_plugins_sheet.column_dimensions['B'].width = 40
    common_plugins_sheet.column_dimensions['C'].width = 20
    
    wb.save(filename=f"plugin list_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx")

def PluginListMaker(vst_path, au_path, vst3_path, aax_path):
    """lists vst, vst3, au and aax plugins"""
    
    vst_list = []
    au_list = []
    vst3_list = []
    aax_list = []


    if vst_path:
        for plugin in os.listdir(vst_path):
            if plugin.endswith('.vst'):
                vst_list.append(plugin)
                continue
            
            new_vst_path = os.path.join(vst_path, plugin)
            if os.path.isdir(new_vst_path):
                for sub_plugin in os.listdir(new_vst_path):
                    if sub_plugin.endswith('.vst'):
                        vst_list.append(sub_plugin)
                    continue
        vst_list.sort()
   
    if au_path:
        for plugin in os.listdir(au_path):
            if plugin.endswith('.component'):
                au_list.append(plugin)
                continue
            
            new_au_path = os.path.join(au_path, plugin)
            if os.path.isdir(new_au_path):
                for sub_plugin in os.listdir(new_au_path):
                    if sub_plugin.endswith('.component'):
                        au_list.append(sub_plugin)
                    continue        
        au_list.sort()
        
    if vst3_path:
        for plugin in os.listdir(vst3_path):
            if plugin.endswith('.vst3'):
                vst3_list.append(plugin)
                continue
            
            new_vst3_path = os.path.join(vst3_path, plugin)
            if os.path.isdir(new_vst3_path):
                for sub_plugin in os.listdir(new_vst3_path):
                    if sub_plugin.endswith('.vst3'):
                        vst3_list.append(sub_plugin)
                    continue
        vst3_list.sort()

    if aax_path:
        for plugin in os.listdir(aax_path):
            if plugin.endswith('.aaxplugin'):
                aax_list.append(plugin)
                continue
            
            new_aax_path = os.path.join(aax_path, plugin)
            if os.path.isdir(new_aax_path):
                for sub_plugin in os.listdir(new_aax_path):
                    if sub_plugin.endswith('.aaxplugin'):
                        aax_list.append(sub_plugin)
                    continue
        aax_list.sort()
        
    ExcelMaker(vst_list, au_list, vst3_list, aax_list)

print("Please enter the plugin directory paths (press enter if no path)")
vst_path = input("Enter the VST plugin directory path: ")
vst3_path = input("Enter the VST3 plugin directory path: ")
au_path = input("Enter the AU plugin directory path: ")
aax_path = input("Enter the AAX plugin directory path: ")

PluginListMaker(vst_path, au_path, vst3_path, aax_path)

print(f"Excel file is created in {os.getcwd()}")
