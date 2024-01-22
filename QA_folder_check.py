# -*- coding: utf-8 -*-
"""
Quality Assesment - site data folder check

@author: Kirill Ivanov
"""


import os
import re
import pip
import sys
import numpy as np
try:
    import pandas as pd
except ImportError:
    pip.main(['install', 'pandas'])
    import pandas as pd
#%% Class
### ----------------------------
class site_folder:
    def __init__(self,master_dir,window):
        self.Mdir = master_dir
        self.window = window
        self.folder_path_check = ['\\1_raw_data\\active_seismic',
                                  '\\1_raw_data\\passive_seismic',
                                  '\\3_site_info\\photos',
                                  '\\3_site_info\\pre_deployment',
                                  '\\3_site_info\\SCS_files',
                                  '\\3_site_info\\field_notes']
        self.column_names = ['SiteName','ArrayID','Active uploaded',
                        'Passive uploaded', 'Pictures uploaded', 'Notes uploaded',
                        'SCS_files', 'Is PreDeployment map moved to data folder?']
        self.flag_table = pd.DataFrame(columns = self.column_names)
        self.QA_check()
    
    def check_data(self):
        self.active, self.passive = ([],[])
        temp_a, temp_b = (0,0)
        for i in range(2):
            data_files = os.listdir(self.subdir_path + self.folder_path_check[i])
            for file in data_files:
                if file.endswith('.dat'):
                    if i == 0:
                        temp_a += 1
                    else: 
                        temp_b += 1
            if i == 0:
                if temp_a > 5:
                    self.active = 'OK'
                else:
                    self.active = 'FLAG'
            else:
                if temp_b < 20:
                    self.passive = 'FLAG'
                else:
                    self.passive = 'OK'
        return self
    
    def check_files(self):
        self.pre_dep, self.scs, self.field_notes = ([],[],[])
        temp_a, temp_b, temp_c = (0,0,0)
        for i in range(3):
            data_files = os.listdir(self.subdir_path + self.folder_path_check[i+3])
            for file in data_files:
                if i == 0:
                    if (re.search('preDeployment', file)) or (re.search('predeployment', file)):
                         temp_a += 1
                elif i == 1:
                    if file.endswith('.log'):
                        temp_b += 1
                else:
                    if re.search('notes', file):
                        temp_c += 1
            if i == 0:
                if temp_a == 0:
                    self.pre_dep = 'FLAG'
                else:
                    self.pre_dep = 'OK'
            elif i == 1:
                if temp_b == 0:
                    self.scs = 'FLAG'
                else:
                    self.scs = 'OK'
            else:
                if temp_c == 0:
                    self.field_notes = 'FLAG'
                else:
                    self.field_notes = 'OK'
        return self
    
    def check_photos(self):
        self.photos = []
        temp = [0,0]
        photo_files = os.listdir(self.subdir_path + self.folder_path_check[2])
        for file in photo_files:
            if (re.search('array_',file))or(re.search('midpoint',file)):
                temp[0] += 1
            elif re.search('HOV_loc', file):
                temp[1] += 1
            else:
                pass
        if (temp[0] == 0) or (temp[1] == 0):
            self.photos = 'FLAG'
        else:
            self.photos = 'OK'
        return self
    
    def clean_table(self):
        # self.flag_table.drop(self.flag_table[all(self.flag_table.iloc[:,2:] == 'OK')].index, 
        #                      inplace=True)
        #self.flag_table.drop(self.flag_table[self.flag_table.iloc[:,2:] == 'OK'].index)
        flag_columns = self.flag_table.columns[2:]
        self.flag_table = self.flag_table[self.flag_table[flag_columns].isin(['FLAG']).any(axis=1)]
        #print(self.flag_table[~self.flag_table[flag_columns].isin(['OK'])])
        return self
    
    def QA_check(self):
        site_dirs = os.listdir(self.Mdir)
        print(f'Number of field site data folder: {len(site_dirs)}')
        self.window.Refresh() if self.window else None
        for subdir in site_dirs:
            try:
                self.subdir_path = os.path.join(self.Mdir,subdir)
                self.ArrayID, self.SiteName = subdir.split('.')
                self.ArrayID = self.ArrayID.replace('_','')
                self.SiteName = self.SiteName.replace('_', ' ')
                print(f'Processing: {self.ArrayID} {self.SiteName}')
                self.window.Refresh() if self.window else None
                self.check_data()
                self.check_photos()
                self.check_files()
                self.flag_table = pd.concat([pd.DataFrame([[self.SiteName,
                                                            self.ArrayID,
                                                            self.active,
                                                            self.passive,
                                                            self.photos,
                                                            self.field_notes,
                                                            self.scs,
                                                            self.pre_dep]],
                                            columns = self.flag_table.columns),
                                             self.flag_table], ignore_index=True)
                
            except (ValueError,FileNotFoundError):
                print(f'{subdir} is not a site data folder.')
                continue
            except PermissionError:
                print(f'Please close Excel file {subdir}')
                continue
        self.clean_table()
        #print(self.flag_table)
        #self.flag_table = self.flag_table.drop([].index)
        #print(self.flag_table.loc[(
        #    )])
#%% Code
### ----------------------------
def main(gui_arg = None, windows = None, index = None):   
    window = windows[0] if windows else None
    final_table = site_folder(os.getcwd(),window).flag_table
    if windows:
        import PySimpleGUI as sg
        #time_input()
        layout_info = [[sg.Table(final_table.values.tolist(),
                                      final_table.columns.tolist(),
                                      auto_size_columns = False,
                                      vertical_scroll_only= False,
                                      def_col_width = 10,
                                      justification='center')]]
        title_info = 'Flag Table'
        index += 1
        new_win = sg.Window(title_info, layout_info,size=(900,300),relative_location=(-500,0)
                            ,finalize=True, resizable=True)
        window.Refresh()
        windows.append(new_win)
        return windows, index
    else:
        pass

    ### ----------------------------
    output_name = 'Quality_Assesment'
    if len(sys.argv) > 1:
        gui_dir = sys.argv[1] +'\\'
    elif gui_arg:
        gui_dir = gui_arg + '\\'
    else:
        gui_dir = ''
    # final_table.style.map(_highlight_flag_cells, subset=final_table.columns[2:])
    #final_table.to_excel(gui_dir + output_name + '.xlsx', index = False)
    # final_table.style.\
    #     apply(_highlight_flag_cells).\
    #     to_excel(gui_dir + output_name + '.xlsx', engine="openpyxl", index = False)
    final_table.style.\
        map(_highlight_flag_cells, props ='color:yellow;background-color:red').\
        to_excel(gui_dir + output_name + '.xlsx',)
#%% Functions
### ----------------------------
    # ----- Information and Test values for scripts ----- #
def _highlight_flag_cells(val,props=''):
    return np.where(val == 'FLAG', props, '')

def time_input():
    return 
#%%
### ----------------------------
# Run this script if it was executed without gui
if __name__ == "__main__":
    main()