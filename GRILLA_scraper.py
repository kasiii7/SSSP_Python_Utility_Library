# -*- coding: utf-8 -*-
"""
Grilla scrapper

@author: Kirill Ivanov
"""


import os
import re
import pandas as pd
import win32com.client 
import sys
import pip
#%%

class GrillaProcessor:
    def __init__(self):
        self.FILE_NAME = pd.DataFrame(columns=['File Name'])
        self.LINE_ID = pd.DataFrame(columns=['LineID', 'LineIDtrace', 'SiteName'])
        self.HEADER = pd.DataFrame(columns=[
            'Instrument', 'Data format', 'Full scale [mV]', 'Start recording', 'End recording',
            'Channel labels', 'Trace length', 'Analyzed', 'Sampling rate', 'Window size',
            'Smoothing type', 'Smoothing', 'GPS location', '(UTC time synchronized to the first recording sample)',
            'Satellite no.', 'Max. H/V'
        ])
        self.SESAME = pd.DataFrame(columns=[
            'f0 > 10 / Lw', 'nc(f0) > 200', 'if  f0 < 0.5Hz',
            'Exists f - in  [f0/4, f0] | AH/V(f -) < A0 / 2',
            'Exists f + in  [f0, 4f0] | AH/V(f +) < A0 / 2', 'A0 > 2 ',
            'fpeak[AH/V(f) ± \uf073A(f)] = f0 ± 5%', '\uf073f < \uf065(f0)',
            '\uf073A(f0) < \uf071(f0)'
        ])
        
    def import_check(self):
        
        
    def _get_lineID_SiteName(self, string):
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
        id_string = string[:string.find('\r')].strip()
        TR_number = id_string[-1]
        id_string_line = id_string.split(',')[0]
        LineID = id_string_line[:10]
        SiteName = id_string_line[10:]
        LineIDtrace = LineID + '.TR' + TR_number
        if LineID.isnumeric() == True:
            return [LineID, LineIDtrace, SiteName]
        else:
            return [id_string_line, id_string_line + '.TR' + TR_number, id_string_line]

    def _read_header(self, string, attr):
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

    def _find_criteria(self, string, trigger):
    
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
            check = re.search('OK',tail[:end])
            if check is not None:
                return check[0]
            else:
                return str('NO')
        else:
            return None

    def process_grilla_files(self):
        grilla_files = []
        cwd = os.getcwd()
        for root, dirs, files in os.walk(cwd):
            for file in files:
                if file.startswith('GRILLA') and file.endswith('.doc'):
                    grilla_files.append(os.path.join(root, file))

        word = win32com.client.Dispatch('Word.Application')
        word.visible = False

        print('Number of GRILLA files found: ' + str(len(grilla_files)))
        for i, path in enumerate(grilla_files):
            file_path_cwd = path.replace(cwd, '')
            print('Working on file: ' + file_path_cwd)

            wb = word.Documents.Open(path)
            doc = word.ActiveDocument
            doc_string = doc.Range().Text
            file_crit = list()

            for num, trig in enumerate(self.SESAME.columns):
                file_crit.append(self._find_criteria(doc_string, trig))
            self.SESAME.loc[i] = file_crit

            file_head_info = list()
            for num, info in enumerate(self.HEADER.columns):
                file_head_info.append(self._read_header(doc_string, info))
            self.HEADER.loc[i] = file_head_info

            self.LINE_ID.loc[i] = self._get_lineID_SiteName(doc_string)
            self.FILE_NAME.loc[i] = file_path_cwd.split('\\')[2].replace('.doc', '').replace('GRILLA', '')

            doc.Close(False)
        
        word.Quit()
        final_table = pd.concat([self.FILE_NAME, self.LINE_ID, self.HEADER, self.SESAME], axis=1)
        output_name = 'GRILLA_all_proccessing'
        if len(sys.argv) > 1:
            gui_dir = sys.argv[1] + '\\'
        else:
            gui_dir = ''
        final_table.to_excel(gui_dir + output_name + '.xlsx', index=False)
        final_table.to_csv(gui_dir + output_name + '.csv', index=False)

if __name__ == "__main__":
    grilla_processor = GrillaProcessor()
    grilla_processor.process_grilla_files()
