# -*- coding: utf-8 -*-
"""
Grilla scrapper

@author: Kirill Ivanov
"""


import os
import re
import pip
import sys
#%% Functions
### ----------------------------

def _get_lineID_SiteName(string, file_path):
    '''
    a function to read the file title and get LineID and Site Name

    Parameters
    ----------
    string : str
        a word grilla file in a string format.

    Returns
    -------
    list : str
        0 - LineID (2023102601)
        1 - LineIDtrace (2023102601).TR1
        2 - SiteName (Washingtion_ES)

    '''
    Filename = file_path.split('\\')[2]
    grilla_number = file_path.split('\\')[3].replace('.doc','')[-1]
    id_string = string[:string.find('\r')].strip()
    TR_number = id_string[-1]
    id_string_line = id_string.split(',')[0]
    LineID = id_string_line[:10]
    SiteName = id_string_line[10:]
    LineIDtrace = LineID + '.TR' + TR_number + '.' + grilla_number
    if LineID.isnumeric() == True:
        return [Filename, LineID, LineIDtrace, SiteName]
    else:
        return [id_string_line, id_string_line + '.TR' + TR_number, id_string_line]
    
def _read_header(string, attr):
    '''
    a fucntion to obtain header information for grilla output file

    Parameters
    ----------
    string : str
        a word grilla file in a string format
    attr : str
        a name of an attribute from the header to get a parameter of

    Returns
    -------
    parameter : str
        an associated parameter of a given attribute

    '''
    m = re.search(re.escape(attr + ':'), string)
    if m is None:
        m = re.search(attr, string)
    # if attr == 'Analyzed':
    #     m = re.search(attr, string)
    if m is not None:
        tail = string[m.span()[1]:]
        if ((attr != 'Start recording') and (attr != 'Trace length')):
            parameter = tail[:tail.find('\r')]
            if attr == 'Instrument':
                parameter = parameter.strip()
            if attr == 'Max. H/V':
                parameter = tail[4:tail.find(' (')]
            return parameter
        else:
            parameter = tail[:tail.find('\t')]
            return parameter
    else:
        return None

def _find_criteria(string, trigger):
    
    '''
    a function to obtain criteria checks for SESAME guidlines (grilla output file)

    Parameters
    ----------
    string : str
        a word grilla file in a string format
    trigger : str
        a trigger for a criteria -- something right before the check.

    Returns
    -------
    ___: str
        OK/NO check

    '''
    
    m = re.search(re.escape(trigger), string)
    if m is not None:
        tail = string[m.span()[1]:]
        end = tail.find('\x07\r')
        if trigger == 'A0 > 2 \r\x07':
            end  = tail[:end].find(' > 2')
            amp_val = tail[:end]
            return amp_val
        check = re.search('OK',tail[:end])
        if check is not None:
            return check[0]
        else:
            return str('NO')
    else:
        return None

#%% Code
### ----------------------------
def main(gui_arg = None, windows = None, index = None, *args):   
    ### checking if Pandas are imported
    ### ----------------------------
    try:
        import pandas as pd
    except ImportError:
        pip.main(['install', 'pandas'])
        import pandas as pd
    
    try:
        import win32com.client 
    except ImportError:
        pip.main(['install', 'pywin32'])
        import win32com.client 
    if windows:
        import PySimpleGUI as sg
        window = windows[0] if windows else None

    ### Header info
    ### ----------------------------
    #FILE_NAME = pd.DataFrame(columns=['File Name'])
    LINE_ID = pd.DataFrame(columns=['File Name','LineID', 'LineIDtrace', 'SiteName'])
    
    ### Header info
    ### ----------------------------
    header_attr_names = ['Instrument',
              'Data format',
              'Full scale [mV]',
              'Start recording',
              'End recording',
              'Channel labels',
              'Trace length',
              'Analyzed',
              'Sampling rate',
              'Window size',
              'Smoothing type',
              'Smoothing',
              'GPS location',
              '(UTC time synchronized to the first recording sample)',
              'Satellite no.',
              'Max. H/V'
        ]
    HEADER = pd.DataFrame(columns=header_attr_names)
    
    ### SESAME guildlines check
    ### ----------------------------
    trigger_look = ['f0 > 10 / Lw',
                    'nc(f0) > 200',
                    'if  f0 < 0.5Hz',
                    'Exists f - in  [f0/4, f0] | AH/V(f -) < A0 / 2',
                    'Exists f + in  [f0, 4f0] | AH/V(f +) < A0 / 2',
                    'A0 > 2 ',
                    'fpeak[AH/V(f) ± \uf073A(f)] = f0 ± 5%',
                    '\uf073f < \uf065(f0)',
                    '\uf073A(f0) < \uf071(f0)',
                    'A0 > 2 \r\x07'
        ]
    SESAME = pd.DataFrame(columns=trigger_look)
    
    ### find all .doc files in directories below 
    ### ----------------------------
    grilla_files = []
    cwd = os.getcwd()
    for root,dirs,files in os.walk(cwd):
        for file in files:
            if (file.startswith('GRILLA'))&(file.endswith('.doc')):
                grilla_files.append(os.path.join(root, file))
    
    
    ### reading each Word .doc file 
    ### IMPORTANT -> close all WORD files 
    ### ----------------------------
    word = win32com.client.Dispatch('Word.Application')
    word.visible = False
    
    ### User Output
    print('Number of GRILLA files found: ' + str(len(grilla_files)))
    window.Refresh() if windows else None
    
    for i,path in enumerate(grilla_files):
        # User Output
        file_path_cwd = path.replace(cwd,'')
        print('Working on file: ' + file_path_cwd)
        if windows:
            sg.one_line_progress_meter('Progress', i+1,len(grilla_files),
                                       orientation = 'h',
                                       no_button = True,
                                       grab_anywhere = True,
                                       keep_on_top = True,
                                       bar_color = ('green','gray'))
            #key = 'OK for 1 meter'
            #meter = sg.QuickMeter.active_meters[key]
            #meter.window.DisableClose = False
            window.Refresh() 
        #open word file
        wb = word.Documents.Open(path)
        doc = word.ActiveDocument
        doc_string = doc.Range().Text
        file_crit = list()
        #get SESAME checks
        for num, trig in enumerate(trigger_look):
            file_crit.append(_find_criteria(doc_string, trig))
        SESAME.loc[i] = file_crit
        #get header info
        file_head_info = list()
        for num, info in enumerate(header_attr_names):
            file_head_info.append(_read_header(doc_string, info))
        HEADER.loc[i] = file_head_info
        #get file name, line and site ID
        LINE_ID.loc[i] = _get_lineID_SiteName(doc_string,file_path_cwd)
        #close word file
        doc.Close(False)    
    #fully quit word 
    word.Quit()
    ### final step, combine line_id, header info, and SESAME checks
    ### ----------------------------
    SESAME.rename(columns={'A0 > 2 \r\x07':'Amplitude'},inplace=True)
    final_table = pd.concat([LINE_ID, HEADER, SESAME], axis=1)
    output_name = 'GRILLA_all_proccessing'
    if len(sys.argv) > 1:
        gui_dir = sys.argv[1] +'\\'
    elif gui_arg:
        gui_dir = gui_arg + '\\'
    else:
        gui_dir = ''
    final_table.to_excel(gui_dir + output_name + '.xlsx', index = False)
    final_table.to_csv(gui_dir + output_name + '.csv', index = False)
    return windows, index
#%%
### ----------------------------
# Run this script if it was executed without gui
if __name__ == "__main__":
    main()