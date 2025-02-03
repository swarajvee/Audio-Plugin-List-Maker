
    """lists vst, vst3, au and aax plugins"""
    
    vst_list = []
    au_list = []
    vst3_list = []
    aax_list = []

    if vst_path:
        vst_list = find_plugins(vst_path, '.vst')
        vst_list.sort()

    if au_path:
        au_list = find_plugins(au_path, '.component')  # Correct file extension for AU plugins
        au_list.sort()

    if vst3_path:
        vst3_list = find_plugins(vst3_path, '.vst3')
        vst3_list.sort()

    if aax_path:
        aax_list = find_plugins(aax_path, '.aaxplugin')
        unused_aax_path = os.path.join(os.path.dirname(aax_path), 'Plug-Ins (Unused)')
        if os.path.isdir(unused_aax_path):
            unused_plugins = find_plugins(unused_aax_path, '.aaxplugin')
            aax_list.extend([plugin.replace('.aaxplugin', ' (Unused).aaxplugin') for plugin in unused_plugins])
        aax_list.sort()

    ExcelMaker(vst_list, au_list, vst3_list, aax_list)

print("Please enter the plugin directory paths (press enter if no path)")
vst_path = input("Enter the VST plugin directory path: ")
vst3_path = input("Enter the VST3 plugin directory path: ")
au_path = input("Enter the AU plugin directory path: ")
aax_path = input("Enter the AAX plugin directory path: ")

PluginListMaker(vst_path, au_path, vst3_path, aax_path)

print(f"Excel file is created in {os.getcwd()}")

