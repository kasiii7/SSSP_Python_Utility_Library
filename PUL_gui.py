# -*- coding: utf-8 -*-
"""
Python Unility Library GUI

@author: Kirill Ivanov
"""

import PySimpleGUI as sg
import subprocess
import os
import grilla_scrapper
import sys
import pandas as pd
   
#%% Functions    

    # ----- Run Imported Scripts ----- #
def run_imported_script(run_directory, out_directory,run_script, window,selected_script):
    '''
    

    Parameters
    ----------
    run_directory : TYPE
        DESCRIPTION.
    out_directory : TYPE
        DESCRIPTION.
    run_script : TYPE
        DESCRIPTION.

    Returns
    -------
    None.

    '''
    try:
        os.chdir(run_directory)
        run_script.main(out_directory, window)
        sg.popup(f"Script '{selected_script}' finished running. \nResults saved in '{out_directory}'")
    except Exception as e:
        sg.popup_error(f"Error running script: {e}")
           
    # ----- Run Any Python Script ----- #
def run_custom_script(run_directory, output_elem, out_directory, window, script_path = None):
    '''
    

    Parameters
    ----------
    run_directory : TYPE
        DESCRIPTION.
    output_elem : TYPE
        DESCRIPTION.
    out_directory : TYPE
        DESCRIPTION.
    window : TYPE
        DESCRIPTION.
    script_path : TYPE, optional
        DESCRIPTION. The default is None.

    Returns
    -------
    None.

    '''
    try:
        os.chdir(run_directory)
        process = subprocess.Popen(['python', script_path, out_directory],
                                   stdout=subprocess.PIPE,
                                   text=True, bufsize=0)
        while True:
            output = process.stdout.readline()
            if output:
                output_elem.print(output.strip())
                window.Refresh() if window else None
            else:
                break
        process.stdout.close()
        process.wait()
        sg.popup(f"Script '{script_path}' finished running. \nResults saved in '{out_directory}'")
    except Exception as e:
        sg.popup_error(f"Error running script: {e}")
        
    # ----- Information and Test values for scripts ----- #
def toolkit_popup(event,table):
    layout_info = [[sg.Text(settings['TOOLKIT'][event+'_info'],justification='left')],
                        [sg.Table([table.values[0].tolist()],
                                  table.columns.tolist(),
                                  auto_size_columns = False, 
                                  vertical_scroll_only= False,
                                  justification='center',
                                  max_col_width=10,
                                  def_col_width = 5,
                                  num_rows=2)],
                        [sg.Exit(button_color='tomato',s=10)]]
    title_info = event + ' Info'
    return sg.Window(title_info, layout_info,size=(1400,230),relative_location=(0,250),
                     finalize=True, resizable=True) 

    
#%% Main
def main_window():
    # ----- Menu ----- #
    menu_def = [['Toolbar', scripts_def],
                ['Help', ['Settings', 'About', 'Exit']]]
    # ----- GUI ----- #
    window_title = settings['GUI']['title']
    scripts_def.append('Select Custom Script')
    layout = [[sg.MenubarCustom(menu_def, tearoff=False)],
              [sg.Text("Select a Script:", justification='r'), sg.Combo(scripts_def, default_value=scripts_def[0], key='-SCRIPTS-',enable_events=True)],
              [sg.Text("Choose a custom script:", justification='r', visible=False, key='-CUSTOM_SCRIPT_TEXT-')],
              [sg.Input(key='-CUSTOM_SCRIPT-', visible=False), sg.FileBrowse(key='-CUSTOM_SCRIPT_BROWSE-', visible=False, file_types=(("Python Files", "*.py"),))],
              [sg.Text("Select Directory Where to Run the Script:", justification='r')],
              [sg.Input(key='-RUN_DIRECTORY-'), sg.FolderBrowse()],
              [sg.Text("Select Directory Where to Save Results:", justification='r')],
              [sg.Input(key='-OUT_DIRECTORY-'), sg.FolderBrowse()],
              [sg.Button("Run Script",s=16, button_color='light green'), sg.Exit(s=10, button_color='tomato')],
              [sg.Text("Output:", justification='r', visible=False, key='-OUTPUT_KEY-'), sg.Output(size=(50, 10), key='-OUTPUT-', visible=False)]]
    # ----- Open Window ----- #
    window = sg.Window(window_title, layout, finalize=True, resizable=True)
    windows = [window]
    index = 0
    # ----- Menu ----- #
    while True:
        #event, values = window.read()
        win, event, values = sg.read_all_windows(timeout=100)
        if event in (sg.WINDOW_CLOSED,"Exit"):
            if win == window:
                break
            win.close()
            windows.remove(win)
            index -= 1
        if any(event in x for x in scripts_def):
            index += 1
            table = pd.DataFrame(data = [settings['TOOLKIT'][event+'_table'].split('//')],
                                      columns = settings['TOOLKIT'][event+'_header'].split('//'))
            new_win = toolkit_popup(event, table)
            windows.append(new_win)
        if event == 'Cancel':
            sg.one_line_progress_meter_cancel()
        if event == 'About':
            sg.popup(window_title, settings['ABOUT']['version'], settings['ABOUT']['description'], settings['ABOUT']['author'], grab_anywhere=True)
        if event == "-SCRIPTS-":
            if values['-SCRIPTS-'] == 'Select Custom Script':
                window['-CUSTOM_SCRIPT-'].update(visible=True)
                window['-CUSTOM_SCRIPT_BROWSE-'].update(visible=True)
                window['-CUSTOM_SCRIPT_TEXT-'].update(visible=True)
            else:
                window['-CUSTOM_SCRIPT-'].update(visible=False)
                window['-CUSTOM_SCRIPT_BROWSE-'].update(visible=False)
                window['-CUSTOM_SCRIPT_TEXT-'].update(visible=False)
        if event == "Run Script":
            window['-OUTPUT_KEY-'].update(visible=True)
            window['-OUTPUT-'].update(visible=True)
            selected_script = values['-SCRIPTS-']
            if selected_script == 'Select Custom Script':
                script_path = values['-CUSTOM_SCRIPT-']
            else:
                script_path = None    
                preselected_scripts = {
                    'GRILLA Scraper': grilla_scrapper
                    }
                run_script = preselected_scripts.get(selected_script)
            run_directory = values['-RUN_DIRECTORY-']
            out_directory = values['-OUT_DIRECTORY-']
            if not out_directory:
                out_directory = gui_directory
            output_elem = window['-OUTPUT-']
            
            if script_path and run_directory:
                output_elem.update('')
                run_custom_script(run_directory, output_elem, out_directory, window,script_path)
                window.refresh()
            elif selected_script and run_directory:
                run_imported_script(run_directory,out_directory,run_script, window,selected_script)
            else:
                sg.popup_error("Please select both script and run_directory!")
    sg.one_line_progress_meter_cancel()
    for win in windows:
        win.close()

#%% Run Script
if __name__ == '__main__':
    gui_directory = os.path.dirname(os.path.abspath(__file__))
    if getattr(sys, 'freeze', False):
        bundle_dir = sys._MEIPASS
    else:
        bundle_dir = gui_directory
    settings = sg.UserSettings(path = bundle_dir, filename='config.ini',use_config_file=True,
                               convert_bools_and_none=True)
    theme = settings['GUI']['theme']
    sg.theme(theme)
    font_family = settings['GUI']['font_family']
    font_size = int(settings['GUI']['font_size'])
    sg.set_options(font=(font_family, font_size))
    scripts_def = settings['SCRIPTS']['scripts'].split(',')
    main_window()





